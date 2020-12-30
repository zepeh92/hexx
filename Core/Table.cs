using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Hexx.Core
{
    public class Table
    {
        private DataTable dataTable;

        /// <summary>
        /// 테이블을 만듭니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="flatSchema">값 만으로 이루어진 스키마.</param>
        public Table(string tableName, Schema flatSchema)
        {
            if (!flatSchema.IsFlat)
            {
                throw new ArgumentException("The schema should be flat");
            }

            dataTable = new DataTable(tableName)
            {
                TableName = tableName,
                CaseSensitive = false
            };

            foreach (Field field in flatSchema.Fields)
            {
                DataColumn column = new DataColumn(field.Name, field.Type.GetAssemblyType())
                {
                    AutoIncrement = field.AutoIncrement,
                    AutoIncrementSeed = field.AutoIncrementSeed,
                    Unique = field.Unique,
                    Caption = field.Description,
                    AllowDBNull = field.Nullable
                };

                if (field.NullDefaultValue != null)
                {
                    column.DefaultValue = field.NullDefaultValue;
                }

                dataTable.Columns.Add(column);
            }

            dataTable.AcceptChanges();

            Schema = flatSchema;
        }

        /// <summary>
        /// 테이블 복제본을 만듭니다.
        /// </summary>
        /// <param name="other"></param>
        public Table(Table other) :
            this(other.Name, other.Schema)
        {
            dataTable = other.dataTable.Clone();
        }

        /// <summary>
        /// 테이블 이름을 반환합니다.
        /// </summary>
        public string Name
        {
            get
            {
                return dataTable.TableName;
            }
            set
            {
                dataTable.TableName = value;
                dataTable.AcceptChanges();
            }
        }

        /// <summary>
        /// 테이블 스키마를 반환합니다.
        /// </summary>
        public Schema Schema
        {
            get;
        }

        /// <summary>
        /// 모든 테이블 행을 반환 합니다.
        /// </summary>
        public IEnumerable<object[]> Rows
        {
            get
            {
                for (int rowIdx = 0; rowIdx != RowCount; ++rowIdx)
                {
                    yield return this[rowIdx];
                }
            }
        }

        /// <summary>
        /// 테이블 행 개수를 반환합니다.
        /// </summary>
        public int RowCount
        {
            get
            {
                return dataTable.Rows.Count;
            }
        }

        /// <summary>
        /// 테이블의 특정 인덱스 행을 반환 합니다.
        /// </summary>
        /// <param name="rowIndex">가져올 행. 행은 0부터 시작 됩니다.</param>
        /// <returns></returns>
        public object[] this[int rowIndex]
        {
            get
            {
                return ConvertDbNullToNull(dataTable.Rows[rowIndex].ItemArray);
            }
        }

        /// <summary>
        /// 다른 테이블과 호환될 수 있는지 여부를 반환합니다.
        /// 서로 호환되는 테이블은 머지가 가능합니다.
        /// </summary>
        /// <param name="other">호환 여부를 검사할 테이블</param>
        /// <returns>True, 호환 가능 시</returns>
        public bool IsCompatibleWith(Table other)
        {
            return Schema.IsCompatibleWith(other.Schema);
        }

        /// <summary>
        /// @todo 제거 예정
        /// 다른 테이블과 행을 머지합니다.
        /// </summary>
        /// <param name="other">머지할 테이블</param>
        public void Merge(Table other)
        {
            if (!IsCompatibleWith(other))
            {
                throw new Exception($"{Name} schema is not compatible with {other.Name} schema");
            }

            AddRows(other.Rows);
        }

        /// <summary>
        /// 행을 하나 추가합니다.
        /// </summary>
        /// <param name="row"></param>
        public void AddRow(object[] row)
        {
            InternalAddRow(row);

            dataTable.AcceptChanges();
        }

        /// <summary>
        /// 행을 하나 추가합니다.
        /// 일치하지 않은 필드 이름이 있다면 무시됩니다.
        /// </summary>
        /// <param name="row"></param>
        public void AddRow((string, object)[] row)
        {
            object[] alignedRow = new object[Schema.FieldCount];

            foreach ((string name, object val) in row)
            {
                int idx = Schema.GetFieldIndex(name);
                if (idx != -1)
                {
                    alignedRow[idx] = val;
                }
            }

            AddRow(alignedRow);
        }

        /// <summary>
        /// 여러 행을 추가합니다.
        /// </summary>
        public void AddRows(IEnumerable<object[]> rows)
        {
            try
            {
                foreach (object[] row in rows)
                {
                    InternalAddRow(row);
                }
            }
            catch(Exception)
            {
                dataTable.RejectChanges();
                throw;
            }

            dataTable.AcceptChanges();
        }

        /// <summary>
        /// 여러행을 추가합니다.
        /// </summary>
        public void AddRows(IEnumerable<(string, object)[]> rows)
        {
            object[] alignedRow = new object[Schema.FieldCount];

            try
            {
                foreach ((string name, object val)[] row in rows)
                {
                    for (int i = 0; i != alignedRow.Length; ++i)
                    {
                        alignedRow[i] = null;
                    }

                    foreach ((string name, object val) in row)
                    {
                        int idx = Schema.GetFieldIndex(name);
                        if (idx != -1)
                        {
                            alignedRow[idx] = val;
                        }
                    }

                    dataTable.Rows.Add(alignedRow);
                }
            }
            catch (Exception)
            {
                dataTable.RejectChanges();
                throw;
            }

            dataTable.AcceptChanges();
        }

        /// <summary>
        /// 모든 행을 제거합니다.
        /// </summary>
        public void ClearRows()
        {
            dataTable.Rows.Clear();
        }

        /// <summary>
        /// 행 인덱스 0부터 특정 값을 지닌 첫 번째 행을 찾습니다.
        /// </summary>
        /// <param name="fieldName">비교할 필드. 없는 필드를 입력할 경우 null이 반환 됩니다.</param>
        /// <param name="value">비교할 필드의 값</param>
        /// <returns>찾은 행을 반환합니다. 없다면 null이 반환 됩니다.</returns>
        public object[] FindFirstRow(string fieldName, object value)
        {
            int columnIndex = Schema.GetFieldIndex(fieldName);
            if (columnIndex == -1)
            {
                return null;
            }

            Field field = Schema[columnIndex];

            // 값과 필드의 타입이 맞지 않다면 맞춰줌
            Type valueAssemblyType = value.GetType();
            Type fieldAssemblyType = field.Type.GetAssemblyType();
            if (valueAssemblyType != fieldAssemblyType)
            {
                value = Convert.ChangeType(value, fieldAssemblyType);
            }

            DataRow[] rows = dataTable.Select(field.Type == FieldType.String ? $"{field.Name} = '{value}'" : $"{field.Name} = {value}");
            if (rows.Length == 0)
            {
                return null;
            }

            return ConvertDbNullToNull(rows.First().ItemArray);
        }

        private void InternalAddRow(object[] row)
        {
            if (Schema.FieldCount < row.Length)
            {// 인자 row의 column이 더 많음. 많은 만큼 잘라서 추가
                dataTable.Rows.Add(row.Take(Schema.FieldCount));
            }
            else
            {
                dataTable.Rows.Add(row);
            }
        }

        /// <summary>
        /// DBNull을 null로 변경해 반환합니다.
        /// </summary>
        private static object[] ConvertDbNullToNull(object[] row)
        {// DataTable에서의 null은 DBNull 타입으로 표현 됨. 그걸 null로 바꿔줌.
            for (int idx = 0; idx != row.Length; ++idx)
            {
                if (row[idx].GetType() == typeof(DBNull))
                {
                    row[idx] = null;
                }
            }
            return row;
        }
    }
}