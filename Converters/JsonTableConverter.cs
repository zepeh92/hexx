using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Unicode;
using System.Linq;
using Hexx.Core;

namespace Hexx.Converters
{
    public class JsonTableConverter : IDisposable
    {
        private static readonly JsonWriterOptions DefaultWriterOption = new JsonWriterOptions()
        {
            Indented = true,
            SkipValidation = false,
            Encoder = JavaScriptEncoder.Create(UnicodeRanges.All)
        };

        public JsonTableConverter()
        {
        }

        ~JsonTableConverter()
        {
        }

        public void Dispose()
        {
        }

        /// <summary>
        /// Value가 null일 때 해당 프로퍼티를 기록할지 여부입니다.
        /// </summary>
        public bool SkipNullValue
        {
            get; set;
        } = true;

        /// <summary>
        /// Container 엘리먼트가 모두 null일 때 해당 프로퍼티를 기록할지 여부입니다.
        /// </summary>
        public bool SkipNullContainer
        {
            get; set;
        } = true;

        /// <summary>
        /// Object 프로퍼티 엘리먼트가 모두 null일 때 해당 프로퍼티를 기록할지 여부입니다.
        /// </summary>
        public bool SkipNullObject
        {
            get; set;
        } = true;

        /// <summary>
        /// Json 포맷의 테이블을 만듭니다.
        /// </summary>
        /// <param name="table">내보낼 테이블</param>
        public void Write(Linker linker, Table table, string outPath)
        {
            MemoryStream stream = new MemoryStream();

            InternalWrite(stream, linker, table);

            File.WriteAllBytes(outPath, stream.ToArray());
        }

        /// <summary>
        /// Json 포맷의 테이블을 만듭니다.
        /// </summary>
        /// <param name="table">내보낼 테이블</param>
        /// <returns></returns>
        public string Write(Linker linker, Table table)
        {
            MemoryStream stream = new MemoryStream();

            InternalWrite(stream, linker, table);

            return Encoding.UTF8.GetString(stream.ToArray());
        }

        /// <summary>
        /// Json 포맷의 테이블로부터 테이블을 만듭니다.
        /// </summary>
        /// <param name="linker"></param>
        /// <param name="schemaName">스키마</param>
        /// <param name="path">Json 포맷 테이블 경로</param>
        /// <returns></returns>
        public Table Read(Linker linker, Schema schema, string path)
        {
            return ReadFromJson(linker, schema, File.ReadAllText(path, Encoding.UTF8));
        }

        /// <summary>
        /// Json 포맷의 테이블로부터 테이블을 만듭니다.
        /// </summary>
        /// <param name="linker"></param>
        /// <param name="schemaName">스키마</param>
        /// <param name="path">Json 포맷 테이블 경로</param>
        /// <returns></returns>
        public Table ReadFromJson(Linker linker, Schema schema, string json)
        {
            JsonDocument doc = JsonDocument.Parse(json);
            Schema flatSchema = linker.ToFlatSchema(schema);
            Dictionary<string, object> keyValues = new Dictionary<string, object>(flatSchema.FieldCount, StringComparer.OrdinalIgnoreCase);
            List<(string, object)> nameAndValues = new List<(string, object)>(flatSchema.FieldCount);
            Table table = new Table(flatSchema.Name, flatSchema);

            foreach (JsonElement jsonRow in doc.RootElement.EnumerateArray())
            {
                GetFieldValues(ref keyValues, jsonRow);

                foreach (var keyValue in keyValues)
                {
                    nameAndValues.Add((keyValue.Key, keyValue.Value));
                }

                table.AddRow(nameAndValues.ToArray());

                nameAndValues.Clear();
                keyValues.Clear();
            }

            return table;
        }

        private void InternalWrite(MemoryStream stream, Linker linker, Table table)
        {
            Schema schema = linker.GetSchema(table.Schema.Name);
            Utf8JsonWriter writer = new Utf8JsonWriter(stream, DefaultWriterOption);

            writer.WriteStartArray();

            foreach (object[] row in table.Rows)
            {
                writer.WriteStartObject();

                int colIdx = 0;

                foreach (Field field in schema.Fields)
                {
                    ParseValue2(writer, linker, table, field, row, ref colIdx);
                }

                writer.WriteEndObject();
            }

            writer.WriteEndArray();
            writer.Flush();
        }

        private void ParseValue2(Utf8JsonWriter writer, Linker linker, Table table, Field field, object[] row, ref int colIdx)
        {
            if (row.Length <= colIdx)
            {
                throw new IndexOutOfRangeException();
            }

            if (field.IsSimpleType)
            {
                if (SkipNullValue)
                {
                    int dummyColIdx = colIdx;
                    if (IsFieldAllNull(linker, field, row, ref dummyColIdx))
                    {
                        SkipField(linker, field, row, ref colIdx);
                        return;
                    }
                }

                object obj = row[colIdx];

                if (string.IsNullOrEmpty(field.Name))
                {
                    if (obj == null)
                    {
                        writer.WriteNullValue();
                    }
                    else
                    {
                        if (field.IsIntegerType)
                        {
                            writer.WriteNumberValue(Convert.ToInt64(obj));
                        }
                        else if (field.IsRealType)
                        {
                            writer.WriteNumberValue(Convert.ToDouble(obj));
                        }
                        else if (field.Type == FieldType.Bool)
                        {
                            writer.WriteBooleanValue(Convert.ToBoolean(obj));
                        }
                        else if (field.Type == FieldType.String)
                        {
                            writer.WriteStringValue(Convert.ToString(obj));
                        }
                        else
                        {
                            throw new Exception($"{table.Schema.Name}.{field.Name} field has an invalid type");
                        }
                    }
                }
                else
                {
                    if (obj == null)
                    {
                        writer.WriteNull(field.Name);
                    }
                    else
                    {
                        if (field.IsIntegerType)
                        {
                            writer.WriteNumber(field.Name, Convert.ToInt64(obj));
                        }
                        else if (field.IsRealType)
                        {
                            writer.WriteNumber(field.Name, Convert.ToDouble(obj));
                        }
                        else if (field.Type == FieldType.Bool)
                        {
                            writer.WriteBoolean(field.Name, Convert.ToBoolean(obj));
                        }
                        else if (field.Type == FieldType.String)
                        {
                            writer.WriteString(field.Name, Convert.ToString(obj));
                        }
                        else
                        {
                            throw new Exception($"{table.Schema.Name}.{field.Name} field has an invalid type");
                        }
                    }
                }

                ++colIdx;
            }
            else if (field.Type == FieldType.Schema)
            {
                if (SkipNullObject)
                {
                    int dummyColIdx = colIdx;
                    if (IsFieldAllNull(linker, field, row, ref dummyColIdx))
                    {
                        SkipField(linker, field, row, ref colIdx);
                        return;
                    }
                }

                if (string.IsNullOrEmpty(field.Name))
                {
                    writer.WriteStartObject();
                }
                else
                {
                    writer.WriteStartObject(field.Name);
                }

                Schema schema = linker.GetSchema(field.RefSchemaName);

                foreach (Field objField in schema.Fields)
                {
                    ParseValue2(writer, linker, table, objField, row, ref colIdx);
                }

                writer.WriteEndObject();
            }
            else if (field.IsContainerType)
            {
                if (SkipNullContainer)
                {
                    int dummyColIdx = colIdx;
                    if (IsFieldAllNull(linker, field, row, ref dummyColIdx))
                    {
                        SkipField(linker, field, row, ref colIdx);
                        return;
                    }
                }

                if (string.IsNullOrEmpty(field.Name))
                {
                    writer.WriteStartArray();
                }
                else
                {
                    writer.WriteStartArray(field.Name);
                }

                foreach (Field elemField in field.Elements)
                {
                    ParseValue2(writer, linker, table, elemField, row, ref colIdx);
                }

                writer.WriteEndArray();
            }
            else
            {
                throw new Exception($"{field.Name} has an invalid type");
            }
        }

        private void SkipField(Linker linker, Field field, object[] row, ref int dummyColIdx)
        {
            if (field.IsSimpleType)
            {
                ++dummyColIdx;
            }
            else if (field.IsContainerType)
            {
                foreach (Field elemField in field.Elements)
                {
                    SkipField(linker, elemField, row, ref dummyColIdx);
                }
            }
            else if (field.Type == FieldType.Schema)
            {
                Schema objSchema = linker.GetSchema(field.RefSchemaName);

                foreach (Field propField in objSchema.Fields)
                {
                    SkipField(linker, propField, row, ref dummyColIdx);
                }
            }
            else
            {
                throw new Exception($"{field.Name} has an invalid type");
            }
        }

        private bool IsFieldAllNull(Linker linker, Field field, object[] row, ref int dummyColIdx)
        {
            if (field.IsSimpleType)
            {
                bool ret = row[dummyColIdx] == null;
                ++dummyColIdx;
                return ret;
            }
            else if (field.IsContainerType)
            {
                foreach (Field elemField in field.Elements)
                {
                    if (!IsFieldAllNull(linker, elemField, row, ref dummyColIdx))
                    {
                        return false;
                    }
                }
            }
            else if (field.Type == FieldType.Schema)
            {
                Schema refSchema = linker.GetSchema(field.RefSchemaName);

                foreach (Field propField in refSchema.Fields)
                {
                    if (!IsFieldAllNull(linker, propField, row, ref dummyColIdx))
                    {
                        return false;
                    }
                }
            }
            else
            {
                throw new Exception($"{field.Name} has an invalid type");
            }

            return true;
        }

        private static void GetFieldValues(ref Dictionary<string, object> keyValues, JsonElement jsonElem, string prefix = null, string postfix = null)
        {
            switch(jsonElem.ValueKind)
            {
                case JsonValueKind.Object:
                    {
                        foreach (JsonProperty jsonProp in jsonElem.EnumerateObject())
                        {
                            string subsetPrefix = prefix == null ? jsonProp.Name : $"{prefix}.{jsonProp.Name}";

                            GetFieldValues(ref keyValues, jsonProp.Value, subsetPrefix, postfix);
                        }
                        break;
                    }
                case JsonValueKind.Array:
                    {
                        int idx = 0;
                        foreach (JsonElement jsonArrElem in jsonElem.EnumerateArray())
                        {
                            if (prefix == null)
                            {
                                GetFieldValues(ref keyValues, jsonArrElem, $"[{idx}]", postfix);
                            }
                            else
                            {
                                GetFieldValues(ref keyValues, jsonArrElem, $"{prefix}[{idx}]", postfix);
                            }
                            ++idx;
                        }
                        break;
                    }
                case JsonValueKind.String:
                    {
                        keyValues.Add($"{prefix}{postfix}", jsonElem.GetString());
                        break;
                    }
                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    {
                        keyValues.Add($"{prefix}{postfix}", null);
                        break;
                    }
                case JsonValueKind.Number: // integer or real
                case JsonValueKind.False: // bool
                case JsonValueKind.True: // bool
                    {
                        keyValues.Add($"{prefix}{postfix}", jsonElem.GetRawText());
                        break;
                    }
                default:
                    {
                        throw new FormatException($"{jsonElem.GetRawText()} is an invalid format text");
                    }
            }
        }
    }
}
