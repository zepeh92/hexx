using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using Hexx.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace Hexx.Converters
{
    public class ExcelTableConverter : IDisposable
    {
        Excel.Application excelApp = null;
        ExcelTableDecorator decorator = new ExcelTableDecorator();

        public ExcelTableConverter()
        {
        }

        ~ExcelTableConverter()
        {
            Dispose();
        }

        public ExcelTableDecorator Decorator
        {
            get
            {
                return decorator;
            }
            set
            {
                if (value == null)
                {
                    decorator = new ExcelTableDecorator();
                }
                else
                {
                    decorator = value;
                }
            }
        }

        public void Dispose()
        {
            if (excelApp != null)
            {
                excelApp.Workbooks.Close();
                if (!excelApp.Quitting)
                {
                    excelApp.Quit();
                    while (excelApp.Quitting)
                    {
                        Thread.Yield();
                    }
                }
                excelApp = null;
            }
        }

        /// <summary>
        /// 엑셀로 부터 테이블을 만듭니다.
        /// </summary>
        /// <param name="schemaName">엑셀 파일의 스키마</param>
        /// <param name="path">엑셀 파일 경로</param>
        /// <returns></returns>
        public Table Read(Linker linker, Schema schema, string path)
        {
            path = Path.GetFullPath(path);

            Excel.Application app = App;

            Table table = new Table(Path.GetFileNameWithoutExtension(path), linker.ToFlatSchema(schema));

            object[,] records = null;
            int rowCount = 0;
            int colCount = 0;
            {
                Excel.Workbook workbook = app.Workbooks.Open(path, ReadOnly: true);
                Excel.Worksheet worksheet = null;
                try
                {
                    Excel.Sheets sheets = workbook.Worksheets;
                    for (int sheetIdx = 1; sheetIdx <= sheets.Count; ++sheetIdx)
                    {
                        Excel.Worksheet sheet = sheets[sheetIdx] as Excel.Worksheet;
                        if (sheet != null && sheet.Name.Equals(table.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            worksheet = sheet;
                        }
                    }

                    if (worksheet != null)
                    {
                        Excel.Range startCell = (Excel.Range)worksheet.Cells[1, 1];
                        Excel.Range endCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                        Excel.Range tableRange = worksheet.Range[startCell, endCell];

                        object[,] tableDatas = tableRange.Value2 as object[,];
                        if (tableDatas != null && tableDatas.Rank >= 2)
                        {// Excel에서 읽어오는 Value2는 1부터 시작함. 이 부분을 Write 때와 맞춰주기 위해 0부터 시작하는 배열로 맞춰줌
                            rowCount = tableDatas.GetLength(0);
                            colCount = tableDatas.GetLength(1);
                            int rowMinIdx = tableDatas.GetLowerBound(0);
                            int colMinIdx = tableDatas.GetLowerBound(1);

                            records = new object[rowCount, colCount];

                            for (int rowIdx = 0; rowIdx < rowCount; ++rowIdx)
                            {
                                for (int colIdx = 0; colIdx < colCount; ++colIdx)
                                {
                                    records[rowIdx, colIdx] = tableDatas[rowIdx + rowMinIdx, colIdx + colMinIdx];
                                }
                            }

                            tableDatas = records;
                            records = null;
                        }

                        decorator.Workbook = workbook;

                        records = decorator.Undecorate(worksheet, schema, tableDatas);
                        if (records == null)
                        {
                            rowCount = 0;
                            colCount = 0;
                        }
                        else
                        {
                            rowCount = records.GetLength(0);
                            colCount = records.GetLength(1);
                        }
                    }
                }
                finally
                {
                    workbook.Close(SaveChanges: false, Filename: path, RouteWorkbook: false);
                }
            }

            for (int rowIdx = 1; rowIdx < rowCount; ++rowIdx)
            {
                (string, object)[] row = new (string, object)[colCount];

                for (int colIdx = 0; colIdx < colCount; ++colIdx)
                {// 필드 이름과 필드 값을 묶어 입력
                    row[colIdx] = ((string)records[0, colIdx], records[rowIdx, colIdx]);
                }

                table.AddRow(row);
            }

            return table;
        }

        /// <summary>
        /// 엑셀 포맷의 테이블을 만듭니다.
        /// </summary>
        /// <param name="table">내보낼 테이블</param>
        /// <param name="outPath">엑셀 테이블 출력 경로</param>
        public void Write(Linker linker, Table table, string outPath)
        {
            outPath = Path.GetFullPath(outPath);

            Excel.Application app = App;

            bool isOverwriting;
            Excel.Workbook workbook;
            Excel.Worksheet worksheet = null;
            if (File.Exists(outPath))
            {// 기존 파일이 있을 경우 덮어 쓰기. 기존 파일의 셀 포맷 유지를 위해.
                isOverwriting = true;
                workbook = app.Workbooks.Open(outPath, ReadOnly: false);
                worksheet = null;

                Excel.Sheets sheets = workbook.Worksheets;
                for (int sheetIdx = 1; sheetIdx <= sheets.Count; ++sheetIdx)
                {
                    Excel.Worksheet sheet = sheets[sheetIdx] as Excel.Worksheet;
                    if (sheet != null && sheet.Name.Equals(table.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = sheet;
                    }
                }

                if (worksheet != null)
                {// 이미 시트가 있다면 시트 컨텐츠만 날림(기존 시트 데코레이션 유지를 위해)
                    Excel.Range usedRange = worksheet.UsedRange;
                    usedRange.ClearContents();
                }
            }
            else
            {
                isOverwriting = false;
                workbook = app.Workbooks.Add();
                if (workbook.Worksheets.Count > 0)
                {// 기본적으로 있는 첫 시트의 이름을(Sheet1) 테이블 이름으로 변경
                    worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                    worksheet.Name = table.Name;
                }
            }

            if (worksheet == null)
            {
                worksheet =
                    workbook.Worksheets.Add(
                        Missing.Value,
                        workbook.Worksheets[workbook.Worksheets.Count], // Add 시 맨 앞에 생성되기 때문에 마지막으로 위치로 시트 이동합니다.
                        Missing.Value,
                        Missing.Value) as Excel.Worksheet;

                worksheet.Name = table.Name;
            }

            // 엑셀 셀 하나하나에 값을 직접 넣으면 속도가 엄청 느림.
            // 그래서 배열에 값들을 매핑하고, 한 번에 엑셀에 밀어넣도록 함.
            object[][] records = table.Rows.ToArray();
            Schema schema = table.Schema;

            object[,] tableDatas = new object[records.Length + 1, schema.FieldCount];

            for (int fieldIdx = 0; fieldIdx < schema.FieldCount; ++fieldIdx)
            {// 프로퍼티 이름
                tableDatas[0, fieldIdx] = schema[fieldIdx].Name;
            }

            for (int rowIdx = 0; rowIdx < records.Length; ++rowIdx)
            {// 레코드들
                for (int colIdx = 0; colIdx < schema.FieldCount; ++colIdx)
                {
                    tableDatas[rowIdx + 1, colIdx] = records[rowIdx][colIdx];
                }
            }

            decorator.Workbook = workbook;

            decorator.Decorate(worksheet, schema, tableDatas);

            if (isOverwriting)
            {
                workbook.Save();
                workbook.Close();
            }
            else
            {
                workbook.SaveAs(outPath, CreateBackup: false);
                workbook.Close();
            }
        }

        private Excel.Application App
        {
            get
            {
                if (excelApp == null)
                {// Excel Application은 생성만 하는데에도 몇 초가 걸림. 
                 // 그래서 로드하도록 첫 엑세스 때 생성하도록 함.
                    excelApp = new Excel.Application();
                }
                return excelApp;
            }
        }
    }
}
