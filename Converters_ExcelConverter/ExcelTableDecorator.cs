using System;
using System.Reflection;
using Hexx.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace Hexx.Converters
{
    public class ExcelTableDecorator
    {
        public Excel.Workbook Workbook { get; set; } = null;

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

        public Excel.Worksheet AddWorksheet(string name)
        {
            Excel.Worksheet newSheet =
                Workbook.Worksheets.Add(
                    Missing.Value,
                    Workbook.Worksheets[Workbook.Worksheets.Count], // Add 시 맨 앞에 생성되기 때문에 마지막으로 위치로 시트 이동합니다.
                    Missing.Value,
                    Missing.Value) as Excel.Worksheet;

            newSheet.Name = name;

            Excel.Worksheet activeSheet = Workbook.ActiveSheet as Excel.Worksheet;
            if (activeSheet != null)
            {
                activeSheet.Select(Type.Missing); // 활성 시트를 기존 시트로
            }

            return newSheet;
        }

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

        public virtual object[,] Undecorate(Excel.Worksheet worksheet, Schema schema, object[,] tableDatas)
        {
            return tableDatas;
        }
    }
}
