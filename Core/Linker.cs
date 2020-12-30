using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Hexx.Core
{
    public class Linker
    {
        private Dictionary<string, Schema> indexedSchemasByName;
        private Dictionary<string, Table> indexedTablesByName;
        private Dictionary<string, List<Table>> indexedPartialTablesBySchemaName;

        public Linker(
            IEnumerable<Schema> schemas, 
            IEnumerable<Table> tables)
        {
            indexedSchemasByName = new Dictionary<string, Schema>(schemas.Count(), StringComparer.OrdinalIgnoreCase);
            indexedTablesByName = new Dictionary<string, Table>(tables.Count(), StringComparer.OrdinalIgnoreCase);
            indexedPartialTablesBySchemaName = new Dictionary<string, List<Table>>(schemas.Count(), StringComparer.OrdinalIgnoreCase);

            foreach (Schema schema in schemas.Concat(from table in tables
                                                     select table.Schema))
            {
                if (indexedSchemasByName.TryGetValue(schema.Name, out Schema existingSchema))
                {// 이미 정의된 스키마가 있음
                    if (!existingSchema.IsCompatibleWith(schema))
                    {
                        throw new Exception($"The schema({schema.Name}) is not compatible with an existing schema");
                    }
                }
                else
                {// 이미 정의된 스키마가 없음
                    indexedSchemasByName.Add(schema.Name, schema);
                    indexedPartialTablesBySchemaName.Add(schema.Name, new List<Table>());
                }
            }

            foreach (Table table in tables)
            {
                Schema schema = table.Schema;

                indexedTablesByName.Add(table.Name, table);

                indexedPartialTablesBySchemaName[schema.Name].Add(table);
            }
        }

        /// <summary>
        /// 모든 스키마를 반환합니다.
        /// </summary>
        public IEnumerable<Schema> Schemas
        {
            get
            {
                return indexedSchemasByName.Values;
            }
        }

        /// <summary>
        /// 모든 테이블을 반환합니다.
        /// </summary>
        public IEnumerable<Table> Tables
        {
            get
            {
                return indexedTablesByName.Values;
            }
        }

        /// <summary>
        /// 스키마를 반환합니다.
        /// </summary>
        /// <param name="name">스키마 이름</param>
        /// <returns>주어진 이름의 스키마를 반환합니다. 없을 시 null이 반환 됩니다.</returns>
        public Schema GetSchema(string name)
        {
            return indexedSchemasByName.TryGetValue(name, out Schema schema) ? schema : null;
        }

        /// <summary>
        /// 스키마 보유 여부를 반환합니다.
        /// </summary>
        /// <param name="name">스키마 이름</param>
        /// <returns>스키마 보유 여부</returns>
        public bool HasSchema(string name)
        {
            return indexedSchemasByName.ContainsKey(name);
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
            return indexedTablesByName.ContainsKey(tableName);
        }

        /// <summary>
        /// 테이블을 반환합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>주어진 이름의 테이블을 반환합니다. 없을 시 null이 반환 됩니다.</returns>
        public Table GetTable(string tableName)
        {
            Table table;
            if (indexedTablesByName.TryGetValue(tableName, out table))
            {
                return table;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 스키마 이름을 기준으로 머지된 테이블을 반환합니다.
        /// 스키마와 파샬 테이블이 모두 없다면 null이 반환 됩니다.
        /// </summary>
        public Table GetMergedTable(string schemaName)
        {
            if (!indexedPartialTablesBySchemaName.ContainsKey(schemaName))
            {
                return null;
            }

            Schema schema = GetSchema(schemaName);
            List<Table> partials = indexedPartialTablesBySchemaName[schemaName];
            
            Table mergedTable = new Table(schemaName, schema);

            foreach (Table partial in partials)
            {
                mergedTable.Merge(partial);
            }

            return mergedTable;
        }

        /// <summary>
        /// 참조하는 값을 반환합니다.
        /// </summary>
        /// <param name="property">필드</param>
        /// <param name="value">값</param>
        /// <returns>참조하는 값. 실패 시 null을 반환합니다.</returns>
        public object GetReferenceValue(string tableSchemaName, string refPropName, object value)
        {
            List<Table> partials;
            if (!indexedPartialTablesBySchemaName.TryGetValue(tableSchemaName, out partials))
            {
                return null;
            }

            if (!partials.Any())
            {
                return null;
            }

            Table firstPartial = partials.First();

            int idx = firstPartial.Schema.GetFieldIndex(refPropName);
            if (idx == -1)
            {
                return null;
            }

            foreach (Table table in partials)
            {
                object[] foundRow = table.FindFirstRow(refPropName, value);
                if (foundRow != null)
                {
                    return foundRow[idx];
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
            if (indexedPartialTablesBySchemaName.TryGetValue(schemaName, out List<Table> partialTables))
            {
                return partialTables;
            }

            return Enumerable.Empty<Table>();
        }

        /// <summary>
        /// 특정 스키마와 연관된 파샬 테이블들을 반환합니다.
        /// </summary>
        /// <returns>연관된 파샬 테이블들</returns>
        public List<Table> GetRelatedPartialTables(string schemaName)
        {
            List<Table> partialTables = new List<Table>();

            foreach (var item in indexedPartialTablesBySchemaName)
            {
                string name = item.Key;
                List<Table> partials = item.Value;

                Schema schema = GetSchema(name);

                foreach (Field prop in schema.Fields)
                {
                    if (IsFieldRelationWithSchemaName(schemaName, prop))
                    {
                        partialTables.AddRange(partials);
                        break;
                    }
                }
            }

            return partialTables;
        }

        /// <summary>
        /// 필드가 특정 스키마와 연관되었는지의 여부를 반환합니다.
        /// </summary>
        /// <param name="schemaName">특정 스키마</param>
        /// <param name="prop">필드</param>
        /// <returns>연관 여부</returns>
        private bool IsFieldRelationWithSchemaName(string schemaName, Field prop)
        {
            if (prop.Type == FieldType.Object)
            {
                if (prop.RefSchemaName.Equals(schemaName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
                else
                {
                    Schema refSchema = GetSchema(prop.RefSchemaName);
                    if (refSchema != null)
                    {
                        foreach (Field refSchemaProp in refSchema.Fields)
                        {
                            if (IsFieldRelationWithSchemaName(schemaName, refSchemaProp))
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            else if (
                prop.Type == FieldType.Ref &&
                prop.RefSchemaName.Equals(schemaName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            else if (prop.IsContainerType)
            {
                return IsFieldRelationWithSchemaName(schemaName, prop.ElementTemplate);
            }
            return false;
        }

        private void InternalRenameRefSchema(Field prop, string schemaName, string newSchemaName)
        {
            if (prop.RefSchemaName.Equals(schemaName))
            {
                prop.RefSchemaName = newSchemaName;
            }

            if (prop.ElementTemplate != null)
            {
                InternalRenameRefSchema(prop.ElementTemplate, schemaName, newSchemaName);
            }
        }

        private void InternalRenameRefProperty(Field prop, string schemaName, string propName, string newPropName)
        {
            if (prop.RefSchemaName.Equals(schemaName))
            {
                if (prop.RefFieldName.Equals(propName))
                {
                    prop.RefFieldName = newPropName;
                }

                if (prop.RefPickedFieldName.Equals(propName))
                {
                    prop.RefPickedFieldName = newPropName;
                }
            }

            if (prop.ElementTemplate != null)
            {
                InternalRenameRefProperty(prop.ElementTemplate, schemaName, propName, newPropName);
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
                    List<Field> flatField = ToFlatFields(field);
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
        private List<Field> ToFlatFields(Field field)
        {
            List<Field> flatFields = new List<Field>();

            switch (field.Type)
            {
                case FieldType.Object:
                    {
                        Schema refSchema = GetSchema(field.RefSchemaName);
                        if (refSchema != null)
                        {
                            return null;
                        }

                        if (refSchema.IsFlat)
                        {
                            flatFields.AddRange(from refSchemaField in refSchema.Fields
                                                select new Field(refSchemaField)
                                                {
                                                    Name = GetObjectFieldName(field, refSchemaField.Name)
                                                });
                        }
                        else
                        {
                            foreach (Field refSchemaProp in refSchema.Fields)
                            {
                                foreach (Field refFlatField in ToFlatFields(refSchemaProp))
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
                                            Name = GetContainerElementName(field, elemIdx++)
                                        };
                        if (field.ElementTemplate.IsSimpleType)
                        {
                            flatFields.AddRange(elemFields);
                        }
                        else
                        {
                            foreach (Field elemField in elemFields)
                            {
                                foreach (Field elemFlatProp in ToFlatFields(elemField))
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
