﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Hexx.Core;

namespace Hexx.Definition
{
    public enum DefinitionType
    {
        Schema,
        EnumTable
    }

    public class DefinitionFile
    {
        Dictionary<string, Table> enumTables = new Dictionary<string, Table>(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, Schema> schemas = new Dictionary<string, Schema>(StringComparer.OrdinalIgnoreCase);
        List<(Schema schema, List<Field> nilFields)> nilFieldSchemas = new List<(Schema schema, List<Field> nilFields)>();
        Dictionary<string, DefinitionType> defs = new Dictionary<string, DefinitionType>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 텍스트로 부터 Definition 만듭니다.
        /// </summary>
        /// <param name="text">텍스트</param>
        /// <returns>Definition 파일</returns>
        public static DefinitionFile Parse(string text)
        {
            JsonSerializerOptions options = new JsonSerializerOptions
            {
                ReadCommentHandling = JsonCommentHandling.Skip,
                Converters =
                {
                    new DefinitionFileConverter()
                }
            };

            return JsonSerializer.Deserialize<DefinitionFile>(text, options);
        }

        /// <summary>
        /// 스키마 그룹 이름
        /// </summary>
        public string Name
        {
            get;
            set;
        } = string.Empty;

        /// <summary>
        /// 스키마들
        /// </summary>
        public IEnumerable<Schema> Schemas
        {
            get
            {
                return schemas.Values;
            }
        }

        /// <summary>
        /// Enum 테이블들
        /// </summary>
        public IEnumerable<Table> EnumTables
        {
            get
            {
                return enumTables.Values;
            }
        }

        /// <summary>
        /// Nil 필드를 가진 스키마 보유 여부
        /// </summary>
        public bool HasNilFieldsSchema
        {
            get
            {
                return nilFieldSchemas.Any();
            }
        }

        public IEnumerable<(Schema, IEnumerable<Field>)> NilFieldSchemas
        {
            get
            {
                foreach ((Schema schema, List<Field> nilFields) in nilFieldSchemas)
                {
                    yield return (schema, nilFields);
                }
            }
        }

        /// <summary>
        /// 이 파일에 정의된 정의 이름 여부를 반환합니다.
        /// </summary>
        /// <param name="name">이름</param>
        /// <returns>True, 정의되었다면</returns>
        public bool Contains(string name)
        {
            return Contains(name, out _);
        }

        /// <summary>
        /// 이 파일에 정의된 정의 이름 여부를 반환합니다.
        /// </summary>
        /// <param name="name">이름</param>
        /// <returns>True, 정의되었다면</returns>
        public bool Contains(string name, out DefinitionType type)
        {
            return defs.TryGetValue(name, out type);
        }

        /// <summary>
        /// 스키마를 추가합니다.
        /// 이미 같은 이름의 스키마가 있을 경우 실패합니다.
        /// </summary>
        /// <param name="newSchema">스키마</param>
        /// <returns>True, 추가 성공 시</returns>
        public bool AddSchema(Schema newSchema)
        {
            if (Contains(newSchema.Name))
            {
                return false;
            }

            schemas.Add(newSchema.Name, newSchema);

            defs.Add(newSchema.Name, DefinitionType.Schema);

            List<Field> nilFields = new List<Field>();

            foreach (Field field in newSchema.Fields)
            {
                if (HasNilField(field))
                {
                    nilFields.Add(field);
                }
            }

            if (nilFields.Any())
            {
                nilFieldSchemas.Add((newSchema, nilFields));
            }

            Weave();
            
            return true;
        }

        /// <summary>
        /// 특정 이름의 스키마를 반환합니다.
        /// 이름은 대소문자를 구분하지 않습니다.
        /// </summary>
        /// <param name="name">스키마 이름</param>
        /// <returns>스키마. 없으면 null이 반환 됩니다.</returns>
        public Schema GetSchema(string name)
        {
            if (!schemas.TryGetValue(name, out Schema schema))
            {
                return null;
            }

            return schema;
        }

        /// <summary>
        /// ENUM 테이블을 추가합니다.
        /// 이미 같은 이름의 ENUM 테이블이 있을 경우 실패합니다.
        /// </summary>
        /// <param name="newEnumTable">ENUM 테이블</param>
        /// <returns>True, 추가 성공 시</returns>
        public bool AddEnumTable(Table newEnumTable)
        {
            if (Contains(newEnumTable.Name))
            {
                return false;
            }

            Field enumNameField = newEnumTable.Schema["name"];
            if (enumNameField == null)
            {
                throw new ArgumentNullException($"name field not found");
            }

            Field enumValueField = newEnumTable.Schema["value"];
            if (enumValueField == null)
            {
                throw new ArgumentNullException($"value field not found");
            }

            enumTables.Add(newEnumTable.Name, newEnumTable);

            defs.Add(newEnumTable.Name, DefinitionType.EnumTable);

            Weave();

            return true;
        }

        /// <summary>
        /// 특정 이름의 ENUM 테이블을 반환합니다.
        /// 이름은 대소문자를 구분하지 않습니다.
        /// </summary>
        /// <param name="name">ENUM 테이블 이름</param>
        /// <returns>ENUM 테이블. 없으면 null이 반환 됩니다.</returns>
        public Table GetEnumTable(string name)
        {
            if (!enumTables.TryGetValue(name, out Table table))
            {
                return null;
            }

            return table;
        }

        /// <summary>
        /// 다른 정의 파일을 참고하여 스키마 필드들의 알 수 없는 타입들의 타입을 정합니다.
        /// </summary>
        /// <param name="files">참고할 다른 정의 파일들</param>
        /// <returns>True, 모든 타입 지정이 완료 되었을 때</returns>
        public void Weave()
        {
            if (!HasNilFieldsSchema)
            {
                return;
            }

            bool nilSchemaHasChanged = false;

            for (int nilSchemaIdx = 0; nilSchemaIdx != nilFieldSchemas.Count; ++nilSchemaIdx)
            {
                (Schema schema, List<Field> nilFields) = nilFieldSchemas[nilSchemaIdx];

                bool nilFieldHasChanged = false;

                for (int nilFieldIdx = 0; nilFieldIdx != nilFields.Count; ++nilFieldIdx)
                {
                    Field nilField = nilFields[nilFieldIdx];

                    if (WeaveField(nilField))
                    {
                        nilFields[nilFieldIdx] = null;
                        nilFieldHasChanged = true;
                    }
                }

                if (nilFieldHasChanged)
                {
                    nilSchemaHasChanged = true;
                    nilFields.RemoveAll(field => field == null);
                }
            }

            if (nilSchemaHasChanged)
            {
                nilFieldSchemas.RemoveAll(v => !v.nilFields.Any());
            }
        }

        /// <summary>
        /// Field 또는 Field의 ElementTemplate으로 Nil FieldType을 가졌는지 재귀적으로 체크합니다.
        /// </summary>
        /// <param name="field">Field</param>
        /// <returns>Nil FieldType 소유 여부</returns>
        private static bool HasNilField(Field field)
        {
            if (field.Type == FieldType.Nil)
            {
                return true;
            }
            else if (field.IsContainerType)
            {
                return HasNilField(field.ElementTemplate);
            }
            else
            {
                return false;
            }
        }

        private bool WeaveField(Field field)
        {
            bool success = true;

            if (field.Type == FieldType.Nil)
            {
                if (defs.TryGetValue(field.TypeName, out DefinitionType defType))
                {
                    switch (defType)
                    {
                        case DefinitionType.Schema:
                            {
                                field.Type = FieldType.Schema;
                            }
                            break;
                        case DefinitionType.EnumTable:
                            {
                                Table enumTable = GetEnumTable(field.TypeName);
                                Field refField = enumTable.Schema["name"];
                                Field repField = enumTable.Schema["value"];
                                if (refField != null && repField != null)
                                {
                                    field.Type = refField.Type;
                                    field.RefFieldName = refField.Name;
                                    field.CachedRefPickedFieldType = repField.Type;
                                    field.RefPickedFieldName = repField.Name;
                                }
                            }
                            break;
                        default:
                            {
                                if (string.IsNullOrEmpty(field.RefTableName))
                                {
                                    success = false;
                                }
                                else
                                {
                                    Schema refTableSchema = GetSchema(field.RefTableName);
                                    Field refField = refTableSchema[field.RefFieldName];
                                    Field repField = refTableSchema[field.RefPickedFieldName];
                                    if (refField != null && repField != null)
                                    {
                                        field.Type = refField.Type;
                                        field.RefFieldName = refField.Name;
                                        field.CachedRefPickedFieldType = repField.Type;
                                        field.RefPickedFieldName = repField.Name;
                                        success = true;
                                    }
                                }
                            }
                            break;
                    }
                }
                else
                {
                    success = false;
                }
            }
            else if (field.IsContainerType)
            {
                if (!WeaveField(field.ElementTemplate))
                {
                    success = false;
                }
            }

            return success;
        }
    }
}
