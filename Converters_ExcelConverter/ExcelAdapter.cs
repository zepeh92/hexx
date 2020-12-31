using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelConverter
{
    public class ExcelAdapter : IDisposable
    {
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook book = null;
        Excel.Worksheet sheet = null;

        public ExcelAdapter()
        {
        }

        ~ExcelAdapter()
        {
            Dispose();
        }

        public Excel.Workbook Workbook
        {
            get
            {
                return book;
            }
        }

        public Excel.Worksheet Worksheet
        {
            get
            {
                return sheet;
            }
        }

        public object[,] WorksheetData
        {
            get
            {
                Excel.Range startCell = (Excel.Range)Worksheet.Cells[1, 1];
                Excel.Range endCell = Worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                Excel.Range range = Worksheet.Range[startCell, endCell];
                object[,] tableDatas = range.Value2 as object[,];

                int rowCount = tableDatas.GetLength(0);
                int colCount = tableDatas.GetLength(1);
                int rowMinIdx = tableDatas.GetLowerBound(0);
                int colMinIdx = tableDatas.GetLowerBound(1);

                object[,] remappedTableDatas = new object[rowCount, colCount];
                for (int rowIdx = 0; rowIdx < rowCount; ++rowIdx)
                {
                    for (int colIdx = 0; colIdx < colCount; ++colIdx)
                    {
                        remappedTableDatas[rowIdx, colIdx] = tableDatas[rowIdx + rowMinIdx, colIdx + colMinIdx];
                    }
                }
                return remappedTableDatas;
            }
            set
            {
                Worksheet.Cells.ClearContents();

                int rowCount = value.GetLength(0);
                int colCount = value.GetLength(1);

                Excel.Range startCell = (Excel.Range)Worksheet.Cells[1, 1];
                Excel.Range endCell = (Excel.Range)Worksheet.Cells[rowCount, colCount];
                Excel.Range range = Worksheet.Range[startCell, endCell];

                range.Value2 = value;
            }
        }

        public void Dispose()
        {
            if (excelApp != null)
            {
                Close();
                excelApp.Workbooks.Close();
                if (!excelApp.Quitting)
                {
                    excelApp.Quit();
                    while (excelApp.Quitting)
                    {
                        Thread.Yield();
                    }
                }
                sheet = null;
                book = null;
                excelApp = null;
            }
        }

        public bool Open(string path, bool readOnly = false)
        {
            path = Path.GetFullPath(path);

            if (!File.Exists(path))
            {
                return false;
            }

            book = excelApp.Workbooks.Open(path, ReadOnly: readOnly);
            if(book.Worksheets.Count > 0)
            {
                sheet = (Excel.Worksheet)book.Worksheets[1];
            }

            return true;
        }

        public bool OpenOrCreate(string path, bool readOnly = false)
        {
            path = Path.GetFullPath(path);

            if (!File.Exists(path))
            {
                return Open(path, readOnly);
            }
            
            Excel.Workbook newBook = excelApp.Workbooks.Add();
            newBook.SaveAs(path, CreateBackup: false);
            newBook.Close();

            return Open(path, readOnly);
        }

        public void Close()
        {
            if (book == null)
            {
                return;
            }
            sheet = null;
            book.Close(SaveChanges: true);
            book = null;
        }

        public IEnumerable<string> Worksheets
        {
            get
            {
                Excel.Sheets sheets = book.Worksheets;
                for (int sheetIdx = 1; sheetIdx <= sheets.Count; ++sheetIdx)
                {
                    Excel.Worksheet sheet = sheets[sheetIdx] as Excel.Worksheet;
                    yield return sheet.Name;
                }
            }
        }

        public string CurrentWorksheet
        {
            get
            {
                if (sheet == null)
                {
                    return null;
                }
                return sheet.Name;
            }
        }

        public bool HasWorksheet(string name)
        {
            return Worksheets.Any(sheetName => sheetName.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        public bool SetWorksheet(string name)
        {
            sheet = GetWorksheet(name);
            return sheet != null;
        }

        public void AddWorksheet(string name)
        {
            Excel.Worksheet newSheet = (Excel.Worksheet)book.Worksheets.Add(Missing.Value, book.Worksheets[book.Worksheets.Count], Missing.Value, Missing.Value);
            newSheet.Name = name;
        }

        public bool RemoveWorksheet(string name)
        {
            Excel.Worksheet targetSheet = GetWorksheet(name);
            if (targetSheet != null)
            {
                targetSheet.Delete();
                return true;
            }
            return false;
        }

        public Excel.Worksheet GetWorksheet(string name)
        {
            Excel.Sheets sheets = book.Worksheets;
            for (int sheetIdx = 1; sheetIdx <= sheets.Count; ++sheetIdx)
            {
                Excel.Worksheet idxSheet = (Excel.Worksheet)sheets[sheetIdx];
                if (idxSheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return idxSheet;
                }
            }
            return null;
        }

        public void SetWorksheetData(object[,] data)
        {
            Excel.Range startRange = (Excel.Range)sheet.Cells[1, 1];
            Excel.Range endRange = sheet.Range[1 + data.GetUpperBound(0), 1 + data.GetUpperBound(1)];
            sheet.Range[startRange, endRange].Value2 = data;
        }

        public object[,] GetWorksheetData()
        {
            Excel.Range startCell = (Excel.Range)sheet.Cells[1, 1];
            Excel.Range endCell = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            return sheet.Range[startCell, endCell].Value2 as object[,];
        }
    }
}
