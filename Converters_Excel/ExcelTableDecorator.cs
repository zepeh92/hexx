using System;
using System.Reflection;
using Hexx.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace Hexx.Converters
{
    public class ExcelTableDecorator
    {
        public Excel.Workbook Workbook { get; set; } = null;

        /// <summary>
        /// 워크시트를 반환합니다.
        /// </summary>
        /// <param name="name">워크시트 이름</param>
        /// <returns>워크시트. 없으면 null이 반환 됩니다.</returns>
        public Excel.Worksheet GetWorksheet(string name)
        {
            Excel.Sheets sheets = Workbook.Worksheets;
            for (int sheetIdx = 1; sheetIdx <= sheets.Count; ++sheetIdx)
            {
                Excel.Worksheet sheet = sheets[sheetIdx] as Excel.Worksheet;
                if (sheet != null && sheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }
            return null;
        }

        /// <summary>
        /// 워크시트를 하나 추가합니다.
        /// </summary>
        /// <param name="name">워크시트 이름</param>
        /// <returns>추가된 워크시트</returns>
        public Excel.Worksheet AddWorksheet(string name)
        {
            Excel.Worksheet prevActiveSheet = Workbook.ActiveSheet as Excel.Worksheet;

            Excel.Worksheet newSheet =
                Workbook.Worksheets.Add(
                    Missing.Value,
                    Workbook.Worksheets[Workbook.Worksheets.Count], // Add 시 맨 앞에 생성되기 때문에 마지막으로 위치로 시트 이동합니다.
                    Missing.Value,
                    Missing.Value) as Excel.Worksheet;

            newSheet.Name = name;

            if (prevActiveSheet != null)
            {
                prevActiveSheet.Select(Type.Missing); // 활성 시트를 기존 시트로
            }

            return newSheet;
        }

        /// <summary>
        /// 워크시트에 테이블 데이터를 적용합니다.
        /// 워크시트를 꾸밀 경우 이 함수를 오버라이드 합니다.
        /// </summary>
        /// <param name="worksheet">워크시트</param>
        /// <param name="schema">테이블 스키마</param>
        /// <param name="tableData">테이블 데이터</param>
        /// <remarks>워크시트에 테이블 데이터를 쓸 때 호출 됨.</remarks>
        public virtual void Decorate(Excel.Worksheet worksheet, Schema schema, object[,] tableData)
        {
            int rowCount = tableData.GetLength(0);
            int colCount = tableData.GetLength(1);
            if (rowCount == 0 || colCount == 0)
            {
                return;
            }

            Excel.Range sheetCells = worksheet.Cells;
            Excel.Range start = (Excel.Range)sheetCells[1, 1];
            Excel.Range end = (Excel.Range)sheetCells[rowCount, colCount];
            Excel.Range range = sheetCells.Range[start, end];

            range.Value2 = tableData;
        }

        /// <summary>
        /// Decorate로 인해 꾸며진 tableData를 ExcelTableConverter가 읽을 수 있는 형태로 변환합니다.
        /// </summary>
        /// <param name="worksheet">테이블 데이터가 있는 워크시트</param>
        /// <param name="schema">테이블 스키마</param>
        /// <param name="tableDatas">테이블 데이터</param>
        /// <returns>ExcelTableConverter에서 인식 가능한 형태의 테이블 데이터</returns>
        /// <remarks>워크시트에서 테이블 데이터를 읽어올 때 호출 됨.</remarks>
        public virtual object[,] Undecorate(Excel.Worksheet worksheet, Schema schema, object[,] tableDatas)
        {
            return tableDatas;
        }
    }
}
