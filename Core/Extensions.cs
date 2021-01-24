using System;

namespace Hexx.Core
{
    public static class Extensions
    {
        /// <summary>
        /// FieldType의 어셈블리 타입을 반환합니다.
        /// </summary>
        public static Type GetAssemblyType(this FieldType type)
        {
            switch (type)
            {
                case FieldType.Bool:
                    return typeof(bool);
                case FieldType.Int8:
                    return typeof(sbyte);
                case FieldType.Int16:
                    return typeof(short);
                case FieldType.Int32:
                    return typeof(int);
                case FieldType.Int64:
                    return typeof(long);
                case FieldType.Real32:
                    return typeof(float);
                case FieldType.Real64:
                    return typeof(double);
                case FieldType.String:
                    return typeof(string);
                default:
                    throw new Exception($"{type} is not a valid field type type");
            }
        }

        /// <summary>
        /// 타입에 따른 기본 값을 반환합니다.
        /// </summary>
        /// <param name="type">필드 타입</param>
        /// <returns>기본 값</returns>
        public static object GetDefaultValue(this FieldType type)
        {
            switch (type)
            {
                case FieldType.Bool:
                    return false;
                case FieldType.Int8:
                    return (sbyte)0;
                case FieldType.Int16:
                    return (short)0;
                case FieldType.Int32:
                    return (int)0;
                case FieldType.Int64:
                    return (long)0;
                case FieldType.Real32:
                    return 0.0f;
                case FieldType.Real64:
                    return 0.0;
                case FieldType.String:
                    return string.Empty;
                case FieldType.List:
                    return null;
                case FieldType.Schema:
                    return null;
                default:
                    return null;
            }
        }
    }
}
