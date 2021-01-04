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

        /// <summary>
        /// 참조용으로 만드는 테이블의 시트 이름
        /// </summary>
        public string RefTableSheetName
        {
            get; set;
        } = "_REF_TABLE";

        public string BooleanTableName
        {
            get; set;
        } = "_BOOLEAN";

        /// <summary>
        /// 행 마진
        /// </summary>
        public int MarginRowCount
        {
            get; set;
        } = 0;

        /// <summary>
        /// 열 마진
        /// </summary>
        public int MarginColumnCount
        {
            get; set;
        } = 0;

        public Excel.XlRgbColor[] ColorPool
        {
            get;
        } = new Excel.XlRgbColor[]
        {
            Excel.XlRgbColor.rgbSkyBlue,
            Excel.XlRgbColor.rgbOrange,
            Excel.XlRgbColor.rgbLimeGreen,
            Excel.XlRgbColor.rgbSandyBrown
        };

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

        /// <summary>
        /// 테이블 데이터를 씁니다.
        /// </summary>
        /// <param name="worksheet">타겟 워크시트</param>
        /// <param name="tableDatas">테이블 데이터</param>
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

            bool hasRef = schema.Fields.Any(prop => prop.Type == FieldType.Ref || !string.IsNullOrEmpty(prop.RefSchemaName) || prop.Type == FieldType.Bool);
            if (!hasRef)
            {
                return;
            }

            Field[] fields = schema.Fields.ToArray();
            Excel.Range sheetCells = worksheet.Cells;

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

            if (fields.Any(field => field.Type == FieldType.Bool))
            {
                Excel.Range startCell = (Excel.Range)enumSheet.Cells[1, 1];
                Excel.Range endCell = (Excel.Range)enumSheet.Cells[3, 1];
                Excel.Range enumRange = enumSheet.Range[startCell, endCell];

                object[,] boolTable = new object[3, 1] 
                {
                    { BooleanTableName },
                    { "TRUE" },
                    { "FALSE" }
                };

                enumRange.Cells.Value2 = boolTable;
                enumRange = enumSheet.Range[startCell.Offset[1], endCell];

                enumRangeAddrs.Add(BooleanTableName, $"={enumSheet.Name}!{enumRange.Address}");
            }
            
            for (int fieldIdx = 0; fieldIdx < fields.Length; ++fieldIdx)
            {// 이 테이블에 사용되는 Enum 셀을 생성 후 주소를 기억
                Field field = fields[fieldIdx];
                if (field.Type != FieldType.Ref && string.IsNullOrEmpty(field.RefSchemaName))
                {// 참조 프로퍼티가 Flat화 되면서 String으로 변경되어서 일단 이렇게 했음. 추후에 방식 수정 필요 
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

                string refTableFieldName;

                if (field.Type == FieldType.Ref || !string.IsNullOrEmpty(field.RefSchemaName))
                {// 참조 프로퍼티가 Flat화 되면서 String으로 변경되어서 일단 이렇게 했음. 추후에 방식 수정 필요 
                    refTableFieldName = $"{field.RefSchemaName}({field.RefFieldName})";
                }
                else if(field.Type == FieldType.Bool)
                {
                    refTableFieldName = BooleanTableName;
                }
                else
                {
                    refTableFieldName = string.Empty;
                }

                if (!enumRangeAddrs.ContainsKey(refTableFieldName))
                {
                    continue;
                }

                string enumRangeAddr = enumRangeAddrs[refTableFieldName];

                Excel.Range startCell = (Excel.Range)sheetCells.Cells[MarginRowCount + 2/* 첫 데이터 행 위치 */, MarginColumnCount + 1 + fieldIdx];
                Excel.Range endCell = (Excel.Range)sheetCells.Cells[MarginRowCount + rowCount, MarginColumnCount + 1 + fieldIdx];
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

            for (int colIdx = 0; colIdx < colCount; ++colIdx)
            {
                Field field = fields[colIdx];
                if (field.Type != FieldType.Ref && string.IsNullOrEmpty(field.RefSchemaName))
                {// 참조 프로퍼티가 Flat화 되면서 String으로 변경되어서 일단 이렇게 했음. 추후에 방식 수정 필요
                    continue;
                }

                // 컬럼에 색칠 되어있을 수 있으니 모두 제거
                Excel.Range colStartCell = (Excel.Range)sheetCells.Cells[MarginRowCount + 1, MarginColumnCount + 1 + colIdx];
                Excel.Range colEndCell = (Excel.Range)sheetCells.Cells[MarginRowCount + rowCount, MarginColumnCount + 1 + colIdx];
                Excel.Range colRange = sheetCells.Range[colStartCell, colEndCell];
                colRange.Interior.ColorIndex = 0;

                for (int rowIdx = 1; rowIdx < rowCount; ++rowIdx)
                {
                    object fieldVal = tableDatas[rowIdx, colIdx];
                    if (fieldVal == null)
                    {
                        continue;
                    }

                    object refVal = linker.GetReferenceValue(field.RefSchemaName, field.RefFieldName, fieldVal);
                    if (refVal == null)
                    {// 값이 없음
                        Excel.Range errCell = (Excel.Range)sheetCells.Cells[MarginRowCount + 1 + rowIdx, MarginColumnCount + 1 + colIdx];
                        errCell.Interior.Color = color;
                    }
                }
            }
        }

        /// <summary>
        /// 필드 행 꾸미기
        /// </summary>
        private void DecorateFieldRow(Excel.Worksheet worksheet, Schema flatSchema)
        {
            if (flatSchema.FieldCount == 0)
            {
                return;
            }

            Excel.Range sheetCells = worksheet.Cells;
            Excel.Range startCell = (Excel.Range)sheetCells[MarginRowCount + 1, MarginColumnCount + 1];
            Excel.Range endCell = (Excel.Range)sheetCells[MarginRowCount + 1, MarginColumnCount + flatSchema.FieldCount];
            Excel.Range range = sheetCells.Range[startCell, endCell];

            // 이름 위치에 컬럼 검색 필터 적용
            range.AutoFilter(1, VisibleDropDown: true);

            // 이름 너비로 컬럼 크기 맞춤
            range.Columns.AutoFit();

            Excel.Borders borders = range.Borders;
            borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            Schema schema = linker.GetSchema(flatSchema.Name);
            if (!schema.IsFlat && ColorPool.Any())
            {// 하나의 필드가 엑셀 상에선 여러 컬럼으로 이루어진다면 색칠한다.
                int colorPoolIdx = 0;
                int cellStartIdx = 1;

                foreach (Field field in schema.Fields)
                {
                    int count = GetFieldCellCount(field);

                    if (count > 1)
                    {
                        Excel.Range fieldStartCell = (Excel.Range)sheetCells[MarginRowCount + 1, MarginColumnCount + cellStartIdx];
                        Excel.Range fieldEndCell = (Excel.Range)sheetCells[MarginRowCount + 1, MarginColumnCount + cellStartIdx + count - 1];
                        Excel.Range fieldRange = sheetCells.Range[fieldStartCell, fieldEndCell];
                        Excel.XlRgbColor color = ColorPool[colorPoolIdx];
                        fieldRange.Interior.Color = color;
                        colorPoolIdx = (colorPoolIdx + 1) % ColorPool.Length;
                    }

                    cellStartIdx = cellStartIdx + count;
                }
            }
        }

        private int GetFieldCellCount(Field field)
        {
            int count = 0;

            if (field.IsContainerType)
            {
                foreach (Field elemField in field.Elements)
                {
                    count += GetFieldCellCount(elemField);
                }
            }
            else if (field.Type == FieldType.Schema)
            {
                Schema schema = linker.GetSchema(field.RefSchemaName);
                
                foreach (Field propField in schema.Fields)
                {
                    count += GetFieldCellCount(propField);
                }
            }
            else
            {
                count = 1;
            }

            return count;
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
