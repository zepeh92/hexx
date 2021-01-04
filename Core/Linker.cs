using System;
using System.Collections.Generic;
using System.Linq;

namespace Hexx.Core
{
    public class Linker
    {
        private readonly Dictionary<string, Schema> indexedSchemas;
        private readonly Dictionary<string, Table> indexedTables;
        private readonly Dictionary<string, List<Table>> indexedPartialTables;
        private readonly Dictionary<string, Table> indexedMergedTables;

        public Linker(IEnumerable<Schema> schemas) :
            this(schemas, Enumerable.Empty<Table>())
        {
        }

        public Linker(
            IEnumerable<Schema> schemas, 
            IEnumerable<Table> tables)
        {
            indexedSchemas = new Dictionary<string, Schema>(schemas.Count(), StringComparer.OrdinalIgnoreCase);
            indexedTables = new Dictionary<string, Table>(tables.Count(), StringComparer.OrdinalIgnoreCase);
            indexedPartialTables = new Dictionary<string, List<Table>>(schemas.Count(), StringComparer.OrdinalIgnoreCase);
            indexedMergedTables = new Dictionary<string, Table>(StringComparer.OrdinalIgnoreCase);

            foreach (Schema schema in schemas.Concat(from table in tables
                                                     select table.Schema))
            {
                if (indexedSchemas.TryGetValue(schema.Name, out Schema existingSchema))
                {// 이미 정의된 스키마가 있음
                    if (!existingSchema.IsCompatibleWith(schema))
                    {
                        throw new Exception($"The schema({schema.Name}) is not compatible with an existing schema");
                    }
                }
                else
                {// 이미 정의된 스키마가 없음
                    indexedSchemas.Add(schema.Name, schema);
                    indexedPartialTables.Add(schema.Name, new List<Table>());
                }
            }

            foreach (Table table in tables)
            {
                Schema schema = table.Schema;

                indexedTables.Add(table.Name, table);

                indexedPartialTables[schema.Name].Add(table);
            }

            foreach (var item in indexedPartialTables)
            {
                Schema schema = GetSchema(item.Key);
                Schema flatSchema = ToFlatSchema(schema);

                Table mergedTable = new Table(schema.Name, flatSchema);

                List<Table> partials = item.Value;
                List<object[]> rows = new List<object[]>(partials.Sum(table => table.RowCount));

                foreach (object[] row in rows)
                {
                    rows.Add(row);
                }

                mergedTable.AddRows(rows);

                indexedMergedTables.Add(mergedTable.Name, mergedTable);
            }
        }

        /// <summary>
        /// 모든 스키마를 반환합니다.
        /// </summary>
        public IEnumerable<Schema> Schemas
        {
            get
            {
                return indexedSchemas.Values;
            }
        }

        /// <summary>
        /// 모든 테이블을 반환합니다.
        /// </summary>
        public IEnumerable<Table> Tables
        {
            get
            {
                return indexedTables.Values;
            }
        }

        /// <summary>
        /// 모든 병합 테이블을 반환합니다.
        /// </summary>
        public IEnumerable<Table> MergedTables
        {
            get
            {
                return indexedMergedTables.Values;
            }
        }

        /// <summary>
        /// 스키마를 반환합니다.
        /// </summary>
        /// <param name="name">스키마 이름</param>
        /// <returns>주어진 이름의 스키마를 반환합니다. 없을 시 null이 반환 됩니다.</returns>
        public Schema GetSchema(string name)
        {
            return indexedSchemas.TryGetValue(name, out Schema schema) ? schema : null;
        }

        /// <summary>
        /// 스키마 보유 여부를 반환합니다.
        /// </summary>
        /// <param name="name">스키마 이름</param>
        /// <returns>스키마 보유 여부</returns>
        public bool HasSchema(string name)
        {
            return indexedSchemas.ContainsKey(name);
        }

        /// <summary>
        /// 임의의 스키마를 플랫 스키마로 변환합니다.
        /// 이미 플랫 스키마일 경우 원본 스키마가 반환 됩니다.
        /// </summary>
        /// <returns>플랫 스키마</returns>
        public Schema ToFlatSchema(Schema schema)
        {
            return schema.IsFlat ? schema : InternalToFlatSchema(schema);
        }

        /// <summary>
        /// 필드가 참조하는 스키마의 필드를 반환합니다.
        /// </summary>
        /// <returns>참조하는 필드. 없을 시 null이 반환 됩니다.</returns>
        public Field GetReferenceField(Field field)
        {
            Schema refSchema = GetSchema(field.RefSchemaName);
            if (refSchema == null)
            {
                return null;
            }

            return refSchema.GetField(field.RefFieldName);
        }

        /// <summary>
        /// 필드가 참조 선택한 스키마의 필드를 반환합니다.
        /// </summary>
        /// <returns>참조 선택한 필드. 없을 시 null이 반환 됩니다.</returns>
        public Field GetReferencePickedField(Field field)
        {
            Schema refSchema = GetSchema(field.RefSchemaName);
            if (refSchema == null)
            {
                return null;
            }

            Field refField = refSchema.GetField(field.RefFieldName);
            if (refField == null)
            {
                return null;
            }

            return refSchema.GetField(refField.RefPickedFieldName);
        }

        /// <summary>
        /// 테이블 보유 여부를 반환합니다.
        /// </summary>
        /// <returns>테이블 보유 여부</returns>
        public bool HasTable(string tableName)
        {
            return indexedTables.ContainsKey(tableName);
        }

        /// <summary>
        /// 테이블을 반환합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>주어진 이름의 테이블을 반환합니다. 없을 시 null이 반환 됩니다.</returns>
        public Table GetTable(string tableName)
        {
            return indexedTables.TryGetValue(tableName, out Table table) ? table : null;
        }

        /// <summary>
        /// 스키마 이름을 기준으로 머지된 테이블을 반환합니다.
        /// 스키마와 파샬 테이블이 모두 없다면 null이 반환 됩니다.
        /// </summary>
        public Table GetMergedTable(string schemaName)
        {
            return indexedMergedTables.TryGetValue(schemaName, out Table table) ? table : null;
        }

        /// <summary>
        /// 참조하는 값을 반환합니다.
        /// </summary>
        /// <param name="property">필드</param>
        /// <param name="value">값</param>
        /// <returns>참조하는 값. 실패 시 null을 반환합니다.</returns>
        public object GetReferenceValue(string tableName, string refFieldName, object value)
        {
            return GetReferenceValue(tableName, refFieldName, refFieldName, value);
        }

        /// <summary>
        /// 참조하는 값을 반환합니다.
        /// </summary>
        /// <param name="property">필드</param>
        /// <param name="value">값</param>
        /// <returns>참조하는 값. 실패 시 null을 반환합니다.</returns>
        public object GetReferenceValue(string tableName, string refFieldName, string refPickedFieldName, object value)
        {
            Table table;
            if (indexedMergedTables.TryGetValue(tableName, out table))
            {
                int refIdx = table.Schema.GetFieldIndex(refFieldName);
                int repIdx = table.Schema.GetFieldIndex(refPickedFieldName);
                if (refIdx != -1 && repIdx != -1)
                {
                    object[] foundRow = table.FindFirstRow(refFieldName, value);
                    if (foundRow != null)
                    {
                        return foundRow[repIdx];
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 특정 스키마를 사용하는 부분 테이블을 반환합니다.
        /// </summary>
        /// <param name="schemaName">스키마 이름</param>
        /// <returns>부분 테이블들</returns>
        public IEnumerable<Table> GetPartialTables(string schemaName)
        {
            if (indexedPartialTables.TryGetValue(schemaName, out List<Table> partialTables))
            {
                return partialTables;
            }

            return Enumerable.Empty<Table>();
        }

        /// <summary>
        /// 특정 스키마와 연관된 파샬 테이블들을 반환합니다.
        /// </summary>
        /// <returns>연관된 파샬 테이블들</returns>
        public IEnumerable<Schema> GetReleatedSchemas(Schema schema)
        {
            HashSet<string> releatedSchemasNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (Field field in schema.Fields)
            {
                InternalGetReleatedSchemaNames(ref releatedSchemasNames, field);
            }

            return releatedSchemasNames.Select(name => GetSchema(name));
        }

        private void InternalGetReleatedSchemaNames(ref HashSet<string> schemaNames, Field field)
        {
            if (field.Type == FieldType.Schema)
            {
                Schema refSchema = GetSchema(field.RefSchemaName);
                if (refSchema != null)
                {
                    schemaNames.Add(refSchema.Name);
                }
            }
            else if (field.Type == FieldType.List)
            {
                InternalGetReleatedSchemaNames(ref schemaNames, field.ElementTemplate);
            }
        }

        /// <summary>
        /// 플랫한 스키마로 만듭니다.
        /// 스키마의 모든 필드는 복제본으로 생성됩니다.
        /// </summary>
        /// <param name="schema">플랫하게 만들 스키마</param>
        /// <returns>플랫해진 스키마. 오류 시 null이 반환 됩니다.</returns>
        private Schema InternalToFlatSchema(Schema schema)
        {
            List<Field> flatFields = new List<Field>();

            foreach (Field field in schema.Fields)
            {
                if (field.IsSimpleType)
                {
                    flatFields.Add(new Field(field));
                }
                else
                {
                    List<Field> flatField = ToFlatFields(field, field.Nullable);
                    if (flatField == null)
                    {
                        return null;
                    }

                    flatFields.AddRange(flatField);
                }
            }

            Schema flatSchema = new Schema()
            {
                Name = schema.Name,
                Description = schema.Description,
                Embed = schema.Embed
            };
            
            foreach (string partialName in schema.Partials)
            {
                flatSchema.AddPartial(partialName);
            }

            foreach (Field field in flatFields)
            {
                flatSchema.AddField(field);
            }

            return flatSchema;
        }

        /// <summary>
        /// 플랫한 필드로 만듭니다.
        /// 인자로 들어온 원본 필드를 수정하지 않습니다.
        /// 필드는 복제본으로 생성 됩니다.
        /// 내부 오류 발생 시 null이 반환 됩니다.
        /// </summary>
        /// <param name="field">플랫하게 만들 필드</param>
        /// <returns>플랫해진 필드. 내부 오류 발생 시 null이 반환 됩니다.</returns>
        private List<Field> ToFlatFields(Field field, bool nullable)
        {
            List<Field> flatFields = new List<Field>();

            switch (field.Type)
            {
                case FieldType.Schema:
                    {
                        Schema refSchema = GetSchema(field.RefSchemaName);
                        if (refSchema == null)
                        {
                            return null;
                        }

                        if (refSchema.IsFlat)
                        {
                            foreach (Field refField in refSchema.Fields)
                            {
                                Field refField2 = new Field(refField)
                                {
                                    Name = GetObjectFieldName(field, refField.Name),
                                    Nullable = nullable
                                };

                                if (nullable)
                                {
                                    refField2.NullDefaultValue = null;
                                }
                                
                                flatFields.Add(refField2);
                            }
                        }
                        else
                        {
                            foreach (Field refSchemaProp in refSchema.Fields)
                            {
                                foreach (Field refFlatField in ToFlatFields(refSchemaProp, nullable))
                                {
                                    if (refFlatField == null)
                                    {
                                        return null;
                                    }

                                    flatFields.Add(new Field(refFlatField)
                                    {
                                        Name = GetObjectFieldName(field, refFlatField.Name)
                                    });
                                }
                            }
                        }
                        break;
                    }
                case FieldType.List:
                    {
                        int elemIdx = 0;
                        var elemFields = from element in field.Elements
                                        select new Field(element)
                                        {
                                            Name = GetContainerElementName(field, elemIdx++),
                                            Nullable = true,
                                            NullDefaultValue = null
                                        };
                        if (field.ElementTemplate.IsSimpleType)
                        {
                            flatFields.AddRange(elemFields);
                        }
                        else
                        {
                            foreach (Field elemField in elemFields)
                            {
                                foreach (Field elemFlatProp in ToFlatFields(elemField, true))
                                {
                                    if (elemFlatProp == null)
                                    {
                                        return null;
                                    }

                                    flatFields.Add(elemFlatProp);
                                }
                            }
                        }
                        break;
                    }
                case FieldType.Ref:
                    {
                        Schema refSchema = GetSchema(field.RefSchemaName);
                        if (refSchema == null)
                        {
                            return null;
                        }

                        // 래퍼런스는 플랫화 시 자신을 참조하는 필드로 변경
                        Field refField = refSchema.GetField(field.RefFieldName);
                        if (refField == null)
                        {
                            return null;
                        }
                        else if (!refField.IsSimpleType)
                        {// 참조가 가리키는 값은 ValueCategoryType만 가리킬 수 있음
                            return null;
                        }
                        else
                        {
                            Field flatField = new Field(field.Name, refField.Type);
                            flatField.NullDefaultValue = field.NullDefaultValue;
                            flatField.RefSchemaName = field.RefSchemaName;
                            flatField.RefFieldName = field.RefFieldName;
                            flatField.RefPickedFieldName = field.RefPickedFieldName;
                            if (field.ElementTemplate != null)
                            {
                                flatField.ElementTemplate = new Field(field.ElementTemplate);
                            }
                            flatField.ElementCount = field.ElementCount;

                            flatFields.Add(flatField);
                        }
                        break;
                    }
                default:
                    {
                        if (field.IsSimpleType)
                        {
                            flatFields.Add(new Field(field));
                        }
                        else
                        {
                            return null;
                        }
                        break;
                    }
            }

            return flatFields;
        }

        private static string GetContainerElementName(Field field, int index)
        {
            return $"{field.Name}[{index}]";
        }

        private static string GetObjectFieldName(Field field, string fieldName)
        {
            return $"{field.Name}.{fieldName}";
        }
    }
}
