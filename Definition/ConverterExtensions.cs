using System.Collections.Generic;
using System.Text.Json;

namespace Hexx.Definition
{
    internal static class ConverterExtensions
    {
        internal static List<string> ReadStringList(this ref Utf8JsonReader reader)
        {
            List<string> items = new List<string>();
            if (reader.TokenType == JsonTokenType.StartArray)
            {
                reader.Read();
                while (reader.TokenType != JsonTokenType.EndArray)
                {
                    items.Add(reader.GetString());
                    reader.Read();
                }
                reader.Read();
            }
            return items;
        }

        /// <summary>
        /// 양 측의 문자열을 감싸는 문자를 제거합니다.
        /// </summary>
        /// <param name="str">문자열</param>
        /// <param name="coverStr">감싸는 문자열</param>
        /// <returns>감싸는 문자열이 제거된 문자열</returns>
        internal static string EraseCover(this string str, string coverStr)
        {
            if (str.StartsWith(coverStr) && str.EndsWith(coverStr) && 
                str.Length >= (coverStr.Length*2))
            {
                return str[coverStr.Length..^coverStr.Length];
            }
            return str;
        }
    }
}
