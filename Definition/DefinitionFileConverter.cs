using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Hexx.Core;

namespace Hexx.Definition
{
    internal class DefinitionFileConverter : JsonConverter<DefinitionFile>
    {
        public readonly static Regex IntTypeRegex = new Regex(@"^int8|int16|int32|int64", RegexOptions.Compiled);
        public readonly static Regex RealTypeRegex = new Regex(@"^float|double|real32|real64", RegexOptions.Compiled);
        public readonly static Regex BoolTypeRegex = new Regex(@"^bool", RegexOptions.Compiled);
        public readonly static Regex ContainerTypeRegex = new Regex(@"^(?<containerType>list)\s*<\s*(?<elementType>.+)\s*,\s*(?<count>\d+)\s*>", RegexOptions.Compiled);
        public readonly static Regex RefTypeRegex = new Regex(@"^(?<refTableName>[\w_]+)\.(?<refPropName>[\w_]+)\.\$(?<repPropName>[\w_]+)", RegexOptions.Compiled);
        public readonly static Regex StringTypeRegex = new Regex(@"^string", RegexOptions.Compiled);
        public readonly static Regex UndefinedTypeRegex = new Regex(@"^[a-zA-Z][\w_]+", RegexOptions.Compiled);
        public readonly static Regex NameRegex = UndefinedTypeRegex;
        public readonly static Regex PartialNameRegex = new Regex(@"^[a-zA-Z][\w\s_/]+", RegexOptions.Compiled);
        public readonly static Regex SimpleFieldDefaultValueAndDescRegex = new Regex(@"(?<defaultValue>[^(//)]+)?(\s*//\s*(?<desc>.+))?");
        public readonly static Regex EnumDefRegex = new Regex(@"^enum (?<name>[a-zA-Z][\w_]+)(\s*:\s*(?<type>int8|int16|int32|int64))?", RegexOptions.Compiled);
        public readonly static Regex EnumRowRegex = new Regex(@"^(?<name>[a-zA-Z_]+)(\s*=\s*(?<value>\d+))?", RegexOptions.Compiled);

        public override bool CanConvert(Type typeToConvert)
        {
            return typeToConvert == typeof(DefinitionFile);
        }

        /// <summary>
        /// 정의 텍스트로 부터 DefinitionFile 인스턴스를 만듭니다.
        /// </summary>
        public override DefinitionFile Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            DefinitionFile defFile = new DefinitionFile();

            if (reader.TokenType == JsonTokenType.StartArray)
            {
                reader.Read();
                while (reader.TokenType != JsonTokenType.EndArray)
                {
                    ReadDefinition(ref reader, ref defFile);
                }
                reader.Read();
            }
            else if (reader.TokenType == JsonTokenType.StartObject)
            {
                ReadDefinition(ref reader, ref defFile);
            }
  
            return defFile;
        }

        public override void Write(Utf8JsonWriter writer, DefinitionFile value, JsonSerializerOptions options)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 정의를 하나 읽습니다.
        /// 정의는 스키마 또는 ENUM 테이블 정의입니다.
        /// </summary>
        /// <param name="reader">reader</param>
        /// <param name="defFile">정의 파일</param>
        private static void ReadDefinition(ref Utf8JsonReader reader, ref DefinitionFile defFile)
        {
            if (reader.TokenType != JsonTokenType.StartObject)
            {
                throw new Exception("An internal error has occured. Definition token must starts a StartObject token");
            }

            if (!reader.Read() && 
                reader.TokenType != JsonTokenType.PropertyName)
            {
                throw new ArgumentException("An invalid definition has detected");
            }

            string firstPropName = reader.GetString();

            if (firstPropName.StartsWith("enum "))
            {
                reader.Read();

                Match match = EnumDefRegex.Match(firstPropName);
                if (!match.Success)
                {
                    throw new ArgumentException($"{firstPropName} is an invalid enum definition");
                }

                GroupCollection groups = match.Groups;
                
                string enumName = groups["name"].Value.Trim();
                if (!NameRegex.IsMatch(enumName))
                {
                    throw MakeDisallowedNameException(enumName, "enum name");
                }

                FieldType enumValueType = FieldType.Nil;
                if (groups["type"].Success)
                {
                    Field valueField = new Field();
                    
                    BuildFieldType(ref valueField, groups["type"].Value);

                    enumValueType = valueField.Type;
                }

                string enumDesc = string.Empty;
                if (firstPropName.IndexOf("//", match.Length) != -1)
                {
                    enumDesc = firstPropName.Substring(firstPropName.IndexOf("//") + 2).Trim();
                }

                Table enumTable = ReadEnumTable(ref reader, enumName, enumValueType);
                if (!string.IsNullOrWhiteSpace(enumDesc))
                {
                    enumTable.Schema.Description = enumDesc;
                }
                
                if (!defFile.AddEnumTable(enumTable))
                {
                    throw new ArgumentException($"{enumName} is an already defined definition");
                }
            }
            else
            {
                Schema schema = new Schema();

                while (reader.TokenType != JsonTokenType.EndObject)
                {
                    string propName = reader.GetString();
                    reader.Read();

                    switch (propName)
                    {
                        case "name":
                            {
                                schema.Name = reader.GetString().Trim();
                                if (!NameRegex.IsMatch(schema.Name))
                                {
                                    throw MakeDisallowedNameException(schema.Name, "schema name");
                                }
                                reader.Read();
                                break;
                            }
                        case "fields":
                            {
                                if (reader.TokenType == JsonTokenType.StartObject)
                                {
                                    reader.Read();
                                    while (reader.TokenType != JsonTokenType.EndObject)
                                    {
                                        string name = reader.GetString().Trim();
                                        if (!NameRegex.IsMatch(name))
                                        {
                                            throw MakeDisallowedNameException(schema.Name, "field name");
                                        }

                                        reader.Read();

                                        Field field = ReadField(ref reader, name);
                                        if (!schema.AddField(field))
                                        {
                                            throw new ArgumentException($"{name} is an already defined field name");
                                        }
                                    }
                                    reader.Read();
                                }
                                break;
                            }
                        case "partials":
                        case "parts":
                            {
                                foreach (string name in reader.ReadStringList())
                                {
                                    if (!PartialNameRegex.IsMatch(name))
                                    {
                                        throw MakeDisallowedNameException(schema.Name, "partial name");
                                    }

                                    if (!schema.AddPartial(name))
                                    {
                                        throw new ArgumentException($"{name} is an already defiend partial name");
                                    }
                                }
                                break;
                            }
                        case "embed":
                            {
                                schema.Embed = reader.GetBoolean();
                                reader.Read();
                                break;
                            }
                        case "desc":
                        case "description":
                            {
                                schema.Description = reader.GetString();
                                reader.Read();
                                break;
                            }
                        default:
                            {
                                reader.Skip();
                                // Skip을 하면 StartToken에서 EndToken으로 가게 되므로,
                                // 한번 더 Read해서 Token을 완전히 끝냄
                                reader.Read();
                                break;
                            }
                    }
                }

                if (!defFile.AddSchema(schema))
                {
                    throw new Exception($"{schema.Name} is an already defiend schema name");
                }
            }

            reader.Read();
        }

        /// <summary>
        /// reader로 부터 ENUM 테이블 행을 읽습니다.
        /// </summary>
        /// <param name="reader">reader</param>
        /// <param name="enumName">ENUM 테이블 이름</param>
        /// <param name="valueType">값 타입. 명시적인 지정이 없다면 Nil을 지정합니다.</param>
        /// <returns>ENUM 테이블</returns>
        private static Table ReadEnumTable(ref Utf8JsonReader reader, string enumName, FieldType valueType)
        {
            var rows = new List<(string name, long value, string comment)>();

            if (reader.TokenType == JsonTokenType.StartArray)
            {
                reader.Read();

                long maxValue = 0;
                long nextValue = 0;

                while (reader.TokenType != JsonTokenType.EndArray)
                {
                    string enumLine = reader.GetString().Trim();
                    reader.Read();

                    Match match = EnumRowRegex.Match(enumLine);
                    if (!match.Success)
                    {
                        throw new ArgumentException($"{enumLine} is an invalid enum row syntax.");
                    }

                    string comment = string.Empty;
                    if (enumLine.IndexOf("//", match.Value.Length) != -1)
                    {
                        comment = enumLine[(enumLine.IndexOf("//", match.Value.Length) + 2)..].Trim();
                    }

                    GroupCollection groups = match.Groups;

                    long value = nextValue;

                    if (groups["value"].Success)
                    {
                        nextValue = value = long.Parse(groups["value"].Value);
                    }

                    ++nextValue;

                    if (value < 0)
                    {
                        throw new OverflowException($"The value({value}) too small. it must be greater than 0");
                    }

                    maxValue = Math.Max(maxValue, value);

                    rows.Add((groups["name"].Value, value, comment));
                }

                if (valueType == FieldType.Nil)
                {
                    if (maxValue <= sbyte.MaxValue)
                    {
                        valueType = FieldType.Int8;
                    }
                    else if (maxValue <= short.MaxValue)
                    {
                        valueType = FieldType.Int16;
                    }
                    else if (maxValue <= int.MaxValue)
                    {
                        valueType = FieldType.Int32;
                    }
                    else if (maxValue <= long.MaxValue)
                    {
                        valueType = FieldType.Int64;
                    }
                    else
                    {
                        throw new OverflowException($"The value({maxValue}) too big. it must be less than {long.MaxValue}");
                    }
                }
                else
                {
                    long typeMaxValue;
                    switch (valueType)
                    {
                        case FieldType.Int8:
                            typeMaxValue = sbyte.MaxValue;
                            break;
                        case FieldType.Int16:
                            typeMaxValue = short.MaxValue;
                            break;
                        case FieldType.Int32:
                            typeMaxValue = int.MaxValue;
                            break;
                        case FieldType.Int64:
                            typeMaxValue = long.MaxValue;
                            break;
                        default:
                            typeMaxValue = 0;
                            break;
                    }

                    if (typeMaxValue < maxValue)
                    {
                        throw new OverflowException($"The value({maxValue}) too big. it must be less than {typeMaxValue}");
                    }
                }

                reader.Read();
            }

            Schema schema = new Schema(enumName, new[]
            {
                    new Field("name", FieldType.String),
                    new Field("value", valueType),
                    new Field("comment", FieldType.String)
            });

            Table table = new Table(enumName, schema);

            foreach (var (name, value, comment) in rows)
            {
                table.AddRow(new object[] { name, value, comment });
            }

            return table;
        }

        private static Field ReadField(ref Utf8JsonReader reader, string name)
        {
            Field field = new Field()
            {
                Name = name
            };

            // Property가 Null 일 때의 기본 값. 이 값은 필드 초기화가 완료된 이후 설정 됨.
            string nullDefaultValueStr = null;

            if (reader.TokenType == JsonTokenType.String)
            {// 간략화 된 필드 정의에서 파싱
                string remainingToken = reader.GetString().Trim();

                reader.Read();

                remainingToken = BuildFieldType(ref field, remainingToken);

                Match match = SimpleFieldDefaultValueAndDescRegex.Match(remainingToken);

                if (match.Groups["defaultValue"].Success)
                {
                    nullDefaultValueStr = match.Groups["defaultValue"].Value.Trim();

                    nullDefaultValueStr = nullDefaultValueStr.EraseCover("'");
                }

                if (match.Groups["desc"].Success)
                {
                    field.Description = match.Groups["desc"].Value;
                }
            }
            else if (reader.TokenType == JsonTokenType.StartObject)
            {
                reader.Read();

                while (reader.TokenType != JsonTokenType.EndObject)
                {
                    string propName = reader.GetString();
                    reader.Read();

                    switch (propName)
                    {
                        case "type":
                            string remainingTypeStr = BuildFieldType(ref field, reader.GetString());
                            if (!string.IsNullOrWhiteSpace(remainingTypeStr))
                            {
                                throw new ArgumentException($"{remainingTypeStr} is an invalid type syntax");
                            }
                            reader.Read();
                            break;
                        case "default":
                            nullDefaultValueStr = reader.GetString();
                            reader.Read();
                            break;
                        case "desc":
                        case "description":
                            field.Description = reader.GetString();
                            reader.Read();
                            break;
                        case "disables":
                            foreach (string tag in reader.ReadStringList())
                            {
                                if (!field.DisableTags.Add(tag.Trim()))
                                {
                                    throw new ArgumentException($"{tag} in an already defined disable tag");
                                }
                            }
                            break;
                        case "groups":
                            foreach (string groupName in reader.ReadStringList())
                            {
                                if (!field.Groups.Add(groupName.Trim()))
                                {
                                    throw new ArgumentException($"{groupName} in an already defined group name");
                                }
                            }
                            break;
                        case "keys":
                            foreach (string keyName in reader.ReadStringList())
                            {
                                if (!field.CompositKeys.Add(keyName.Trim()))
                                {
                                    throw new ArgumentException($"{keyName} in an already defined key name");
                                }
                            }
                            break;
                        case "nullable":
                            field.Nullable = reader.GetBoolean();
                            reader.Read();
                            break;
                        case "auto_increment":
                            field.AutoIncrement = reader.GetBoolean();
                            reader.Read();
                            break;
                        case "unique":
                            field.Unique = reader.GetBoolean();
                            reader.Read();
                            break;
                        case "non_serialized":
                            field.NonSerialized = reader.GetBoolean();
                            reader.Read();
                            break;
                        default:
                            reader.Skip();
                            // Skip을 하면 StartToken에서 EndToken으로 가게 되므로,
                            // 한번 더 Read해서 Token을 완전히 끝냄
                            reader.Read();
                            break;
                    }
                }

                reader.Read();
            }
            else
            {
                throw new ArgumentException($"An invalid schema fields definition has detected. should define a string or object in the schema field");
            }

            if (nullDefaultValueStr != null)
            {
                field.NullDefaultValue = nullDefaultValueStr;
            }
            
            return field;
        }

        /// <summary>
        /// 문자열로 부터 필드 타입을 재귀적으로 빌드합니다.
        /// </summary>
        /// <param name="field">필드</param>
        /// <param name="str">문자열</param>
        /// <returns>빌드 후 남은 문자열</returns>
        private static string BuildFieldType(ref Field field, string str)
        {
            Match match;
            if (IntTypeRegex.IsMatch(str))
            {
                match = IntTypeRegex.Match(str);
                
                field.TypeName = match.Value;
                switch (match.Value)
                {
                    case "int8":
                        field.Type = FieldType.Int8;
                        break;
                    case "int16":
                        field.Type = FieldType.Int16;
                        break;
                    case "int32":
                        field.Type = FieldType.Int32;
                        break;
                    case "int64":
                        field.Type = FieldType.Int64;
                        break;
                    default:
                        throw new Exception($"An internal error has occured. {match.Value} is a unknown integer type syntax.");
                }
            }
            else if (RealTypeRegex.IsMatch(str))
            {
                match = RealTypeRegex.Match(str);
                
                field.TypeName = match.Value;
                switch (str)
                {
                    case "float":
                    case "real32":
                        field.Type = FieldType.Real32;
                        break;
                    case "double":
                    case "real64":
                        field.Type = FieldType.Real64;
                        break;
                    default:
                        throw new Exception($"An internal error has occured. {match.Value} is a unknown real type syntax.");
                }
            }
            else if (BoolTypeRegex.IsMatch(str))
            {
                match = BoolTypeRegex.Match(str);
                
                field.Type = FieldType.Bool;
                field.TypeName = match.Value;
            }
            else if (StringTypeRegex.IsMatch(str))
            {
                match = StringTypeRegex.Match(str);

                field.Type = FieldType.String;
                field.TypeName = match.Value;
            }
            else if (ContainerTypeRegex.IsMatch(str))
            {
                match = ContainerTypeRegex.Match(str);
                GroupCollection groups = match.Groups;
                
                field.Type = FieldType.List;
                field.TypeName = match.Value;
                field.ElementCount = int.Parse(groups["count"].Value);

                Field elemField = new Field();

                BuildFieldType(ref elemField, groups["elementType"].Value);

                field.ElementTemplate = elemField;
            }
            else if (RefTypeRegex.IsMatch(str))
            {
                match = RefTypeRegex.Match(str);
                GroupCollection groups = match.Groups;

                field.Type = FieldType.Ref;
                field.TypeName = str;
                field.RefTableName = groups["refTableName"].Value;
                field.RefFieldName = groups["refPropName"].Value;
                field.RefPickedFieldName = groups["repPropName"].Value;
            }
            else if (UndefinedTypeRegex.IsMatch(str))
            {
                match = UndefinedTypeRegex.Match(str);

                field.Type = FieldType.Nil;
                field.TypeName = match.Value;
                field.RefSchemaName = match.Value;
            }
            else
            {
                throw new ArgumentException($"{str} is an invalid field type syntax");
            }

            return str[(match.Value.Length + match.Index)..].Trim();
        }

        private static ArgumentException MakeDisallowedNameException(string name, string who)
        {
            return new ArgumentException($"{name} is a disallowed {who}");
        }
    }
}
