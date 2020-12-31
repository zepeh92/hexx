using System;
using System.Collections.Generic;
using System.Linq;
using Hexx.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace Hexx.Converters
{
    public class BrigriExcelTableDecorator : ExcelTableDecorator
    {
        Linker linker;

        public BrigriExcelTableDecorator(Linker linker)
        {
            this.linker = linker;
        }

        public bool MarkOnRefError
        {
            get; set;
        } = true;

        public string RefTableSheetName
        {
            get; set;
        } = "_REF_TABLE";

        public int MarginRowCount
        {
            get; set;
        } = 0;

        public int MarginColumnCount
        {
            get; set;
        } = 0;

        public override object[,] Undecorate(Excel.Worksheet worksheet, Schema schema, object[,] tableDatas)
        {
            int rowCount = tableDatas.GetLength(0);
            int colCount = tableDatas.GetLength(1);

            if (rowCount <= MarginRowCount || colCount <= MarginColumnCount)
            {
                return null;
            }

            for (int colIdx = MarginColumnCount; colIdx < colCount; ++colIdx)
            {// 필드 이름 컬럼에 공백이나 null이 있다면 거기까지만 필드 컬럼으로
                if (string.IsNullOrWhiteSpace(tableDatas[MarginRowCount, colIdx] as string))
                {
                    colCount = colIdx;
                    break;
                }
            }

            for (int rowIdx = MarginColumnCount; rowIdx < rowCount; ++rowIdx)
            {// 전부 빈 행이 있다면 거기까지만 유효한 행으로 인식
                bool allNull = true;
                for (int colIdx = MarginColumnCount; colIdx < colCount; ++colIdx)
                {
                    if (tableDatas[rowIdx, colIdx] != null)
                    {
                        allNull = false;
                        break;
                    }
                }
                if (allNull)
                {
                    rowCount = rowIdx;
                }
            }

            object[,] remappedTableDatas = new object[rowCount - MarginRowCount, colCount - MarginColumnCount];

            for (int rowIdx = MarginRowCount; rowIdx < rowCount; ++rowIdx)
            {// 마진 떼어내기
                for (int colIdx = MarginColumnCount; colIdx < colCount; ++colIdx)
                {
                    remappedTableDatas[rowIdx - MarginRowCount, colIdx - MarginColumnCount] = tableDatas[rowIdx, colIdx];
                }
            }

            return remappedTableDatas;
        }

        public override void Decorate(Excel.Worksheet worksheet, Schema schema, object[,] tableDatas)
        {
            WriteTableDatas(worksheet, tableDatas);

            ApplyDropboxOnEnumCells(worksheet, schema, tableDatas);

            PaintOnRefErrorCells(worksheet, schema, tableDatas);

            DrawOutline(worksheet, tableDatas);

            DecorateFieldRow(worksheet, schema);
        }

        private void WriteTableDatas(Excel.Worksheet worksheet, object[,] tableDatas)
        {
            int rowCount = tableDatas.GetLength(0);
            int colCount = tableDatas.GetLength(1);
            if (rowCount == 0 && colCount == 0)
            {
                return;
            }

            Excel.Range sheetCells = worksheet.Cells;

            Excel.Range startCell = (Excel.Range)sheetCells.Cells[MarginRowCount + 1, MarginColumnCount + 1];
            Excel.Range endCell = (Excel.Range)sheetCells.Cells[MarginRowCount + rowCount, MarginColumnCount + colCount];
            Excel.Range range = sheetCells.Range[startCell, endCell];

            range.Value2 = tableDatas;
        }

        /// <summary>
        /// ENUM 값에 드롭박스 적용
        /// </summary>
        private void ApplyDropboxOnEnumCells(Excel.Worksheet worksheet, Schema schema, object[,] tableDatas)
        {
            int rowCount = tableDatas.GetLength(0);
            int colCount = tableDatas.GetLength(1);

            if (rowCount <= 1 || colCount == 0)
            {
                return;
            }

            Field[] fields = schema.Fields.ToArray();
            Excel.Range sheetCells = worksheet.Cells;

            bool hasRef = fields.Any(prop => prop.Type == FieldType.Ref);
            if (!hasRef)
            {
                return;
            }

            Excel.Worksheet enumSheet = GetWorksheet(RefTableSheetName);
            if (enumSheet == null)
            {
                enumSheet = AddWorksheet(RefTableSheetName);
            }
            else
            {
                enumSheet.Cells.Clear();
            }

            Dictionary<string, string> enumRangeAddrs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            for (int fieldIdx = 0; fieldIdx < fields.Length; ++fieldIdx)
            {// 이 테이블에 사용되는 Enum 셀을 생성 후 주소를 기억
                Field field = fields[fieldIdx];
                if (field.Type != FieldType.Ref)
                {
                    continue;
                }

                string refTablePropName = $"{field.RefSchemaName}({field.RefFieldName})";

                if (enumRangeAddrs.ContainsKey(refTablePropName))
                {
                    continue;
                }

                if (!linker.HasTable(field.RefSchemaName))
                {
                    continue;
                }

                Table enumTable = linker.GetTable(field.RefSchemaName);

                int refPropIndex = enumTable.Schema.GetFieldIndex(field.RefFieldName);
                if (refPropIndex == -1)
                {
                    continue;
                }

                if (enumTable.RowCount == 0)
                {
                    continue;
                }

                object[,] enumValueNames = new object[enumTable.RowCount + 1, 1];
                enumValueNames[0, 0] = refTablePropName;

                for (int rowIdx = 0; rowIdx != enumTable.RowCount; ++rowIdx)
                {
                    enumValueNames[rowIdx + 1, 0] = enumTable[rowIdx][refPropIndex];
                }

                int enumRangeColumnIdx = 1 + enumRangeAddrs.Count;

                Excel.Range startCell = enumSheet.Cells[1, enumRangeColumnIdx] as Excel.Range;
                Excel.Range endCell = enumSheet.Cells[1 + enumTable.RowCount, enumRangeColumnIdx] as Excel.Range;
                Excel.Range enumRange = enumSheet.Range[startCell, endCell];

                enumRange.Value2 = enumValueNames;

                enumRange = enumSheet.Range[startCell.Offset[1], endCell];

                enumRangeAddrs.Add(refTablePropName, $"={enumSheet.Name}!{enumRange.Address}");
            }

            // 테이블 레코드에 필터로 넣음
            for (int fieldIdx = 0; fieldIdx < fields.Length; ++fieldIdx)
            {
                Field field = fields[fieldIdx];
                if (field.Type != FieldType.Ref)
                {
                    continue;
                }

                string refTableFieldName = $"{field.RefSchemaName}({field.RefFieldName})";

                if (!enumRangeAddrs.ContainsKey(refTableFieldName))
                {
                    continue;
                }

                string enumRangeAddr = enumRangeAddrs[refTableFieldName];

                Excel.Range startCell = (Excel.Range)sheetCells.Cells[MarginRowCount + 2, MarginColumnCount + 1 + fieldIdx];
                Excel.Range endCell = (Excel.Range)sheetCells.Cells[MarginRowCount + 2 + rowCount, MarginColumnCount + 1 + fieldIdx];
                Excel.Range range = sheetCells.Range[startCell, endCell];

                Excel.Validation valid = range.Validation;
                if (valid != null)
                {
                    valid.Delete();
                }
                
                valid.Add(
                        Excel.XlDVType.xlValidateList,
                        Excel.XlDVAlertStyle.xlValidAlertWarning,
                        Excel.XlFormatConditionOperator.xlBetween,
                        enumRangeAddr,
                        Type.Missing);
                valid.InCellDropdown = true;
            }
        }

        /// <summary>
        /// 참조 에러 셀에 색칠
        /// </summary>
        private void PaintOnRefErrorCells(Excel.Worksheet worksheet, Schema schema, object[,] tableDatas, Excel.XlRgbColor color = Excel.XlRgbColor.rgbDarkRed)
        {
            int rowCount = tableDatas.GetLength(0);
            int colCount = tableDatas.GetLength(1);

            if (rowCount <= 1 || colCount == 0)
            {
                return;
            }

            Field[] fields = schema.Fields.ToArray();
            Excel.Range sheetCells = worksheet.Cells;

            for (int rowIdx = 0; rowIdx < rowCount; ++rowIdx)
            {// 레코드 - 참조 오류에 색칠
                for (int colIdx = 0; colIdx < colCount; ++colIdx)
                {
                    object fieldVal = tableDatas[rowIdx, colIdx];
                    if (fieldVal == null)
                    {
                        continue;
                    }

                    Field field = fields[colIdx];
                    if (string.IsNullOrEmpty(field.RefSchemaName))
                    {
                        continue;
                    }

                    object refVal = linker.GetReferenceValue(field.RefSchemaName, field.RefFieldName, fieldVal);
                    if (refVal == null)
                    {// 값이 없음
                        Excel.Range errCell = (Excel.Range)sheetCells.Cells[MarginRowCount + 2 + rowIdx, MarginColumnCount + 1 + colIdx];
                        errCell.Interior.Color = color;
                    }
                }
            }
        }

        /// <summary>
        /// 필드 행 꾸미기
        /// </summary>
        private void DecorateFieldRow(Excel.Worksheet worksheet, Schema schema)
        {
            if (schema.FieldCount == 0)
            {
                return;
            }

            Excel.Range sheetCells = worksheet.Cells;
            Excel.Range startCell = (Excel.Range)sheetCells[MarginRowCount + 1, MarginColumnCount + 1];
            Excel.Range endCell = (Excel.Range)sheetCells[MarginRowCount + 1, MarginColumnCount + schema.FieldCount];
            Excel.Range range = sheetCells.Range[startCell, endCell];

            // 프로퍼티 이름 위치에 컬럼 검색 필터 적용
            range.AutoFilter(1, VisibleDropDown: true);

            // 프로퍼티 이름 너비로 컬럼 크기 맞춤
            range.Columns.AutoFit();

            Excel.Borders borders = range.Borders;
            borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            borders.Weight = Excel.XlBorderWeight.xlMedium;
        }

        /// <summary>
        /// 외곽선 칠하기
        /// </summary>
        private void DrawOutline(Excel.Worksheet worksheet, object[,] tableDatas)
        {
            int colCount = tableDatas.GetLength(1);
            if (colCount == 0)
            {
                return;
            }

            int rowCount = tableDatas.GetLength(0);

            Excel.Range sheetCells = worksheet.Cells;
            Excel.Range startCell = (Excel.Range)sheetCells.Cells[MarginRowCount + 1, MarginColumnCount + 1];
            Excel.Range endCell= (Excel.Range)sheetCells.Cells[MarginRowCount + 1 + rowCount, MarginColumnCount + 1 + colCount];
            Excel.Range range = sheetCells.Range[startCell, endCell];

            Excel.Borders cellBordersRange = range.Borders;
            cellBordersRange.Color = Excel.XlRgbColor.rgbLightGray;
            cellBordersRange.LineStyle = Excel.XlLineStyle.xlContinuous;
        }
    }
}
