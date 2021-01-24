using System;
using System.Collections.Generic;
using System.Linq;

namespace Hexx.Core
{
    public enum FieldType : byte
    {
        Nil = 0,
        Bool,
        Int8,
        Int16,
        Int32,
        Int64,
        Real32,
        Real64,
        String,
        List,
        Schema,
    }

    public class Field
    {
        private object defaultValue = null;
        private Field elementTemplate = null;

        public Field()
        {
        }

        public Field(FieldType type) :
            this(string.Empty, type)
        {
        }

        public Field(string name, FieldType type)
        {
            Name = name;
            Type = type;
        }

        /// <summary>
        /// 필드를 깊은 복사로 생성합니다.
        /// </summary>
        /// <param name="other"></param>
        public Field(Field other)
        {
            Name = other.Name;
            Type = other.Type;
            TypeName = other.TypeName;
            Description = other.Description;
            AutoIncrement = other.AutoIncrement;
            Unique = other.Unique;
            NonSerialized = other.NonSerialized;
            Nullable = other.Nullable;
            NullDefaultValue = other.NullDefaultValue;
            RefTableName = other.RefTableName;
            RefSchemaName = other.RefSchemaName;
            RefFieldName = other.RefFieldName;
            RefPickedFieldName = other.RefPickedFieldName;
            Groups = new HashSet<string>(other.Groups);
            DisableTags = new HashSet<string>(other.DisableTags);
            CompositKeys = new HashSet<string>(other.CompositKeys);
            if (other.ElementTemplate != null)
            {
                ElementTemplate = new Field(other.ElementTemplate);
            }
            ElementCount = other.ElementCount;
        }

        /// <summary>
        /// 필드 이름
        /// </summary>
        public string Name
        {
            get;
            set;
        } = string.Empty;

        /// <summary>
        /// 타입
        /// </summary>
        public FieldType Type
        {
            get;
            set;
        } = FieldType.Nil;

        /// <summary>
        /// 타입 이름
        /// </summary>
        public string TypeName
        {
            get;
            set;
        } = string.Empty;

        /// <summary>
        /// 필드에 대한 설명
        /// </summary>
        public string Description
        {
            get;
            set;
        } = string.Empty;

        /// <summary>
        /// 자동 값 증가 여부
        /// </summary>
        public bool AutoIncrement
        {
            get;
            set;
        } = false;

        /// <summary>
        /// 자동 값 증가 시작 값
        /// </summary>
        public int AutoIncrementSeed
        {
            get;
            set;
        } = 1;

         /// <summary>
        /// 테이블 상에서 이 필드의 값이 고유해야 하는지 여부
        /// </summary>
        public bool Unique
        {
            get;
            set;
        } = false;

        /// <summary>
        /// 직렬화하지 않음 여부
        /// </summary>
        public bool NonSerialized
        {
            get;
            set;
        } = false;

        /// <summary>
        /// null 값 허용 여부
        /// Nullable이 false인 테이블 필드에 null 입력 시 예외가 발생합니다.
        /// </summary>
        public bool Nullable
        {
            get;
            set;
        } = false;

        /// <summary>
        /// null 입력 시 설정될 값 입니다.
        /// 이 값은 Nullable이 아닐 때에만 지정 가능합니다.
        /// </summary>
        public object NullDefaultValue
        {
            get
            {
                return defaultValue;
            }
            set
            {
                if (value == null)
                {
                    defaultValue = null;

                    if (IsContainerType && ElementTemplate != null)
                    {
                        ElementTemplate.NullDefaultValue = null;
                    }
                }
                else
                {
                    switch (Type)
                    {
                        case FieldType.Bool:
                            defaultValue = Convert.ToBoolean(value);
                            break;
                        case FieldType.Int8:
                            defaultValue = Convert.ToSByte(value);
                            break;
                        case FieldType.Int16:
                            defaultValue = Convert.ToInt16(value);
                            break;
                        case FieldType.Int32:
                            defaultValue = Convert.ToInt32(value);
                            break;
                        case FieldType.Int64:
                            defaultValue = Convert.ToInt64(value);
                            break;
                        case FieldType.Real32:
                            defaultValue = Convert.ToSingle(value);
                            break;
                        case FieldType.Real64:
                            defaultValue = Convert.ToDouble(value);
                            break;
                        case FieldType.String:
                            defaultValue = Convert.ToString(value);
                            break;
                        case FieldType.Schema:
                            defaultValue = null;
                            break;
                        //case FieldType.Ref:
                        //    defaultValue = value;
                        //    break;
                        default:
                            if (IsContainerType)
                            {// 컨테이너 타입은 하위 요소에 default value를 전파
                                defaultValue = value;
                                ElementTemplate.NullDefaultValue = value;
                            }
                            else
                            {
                                defaultValue = null;
                            }
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// 스키마를 반환합니다.
        /// 이 값은 OBJECT 타입일 때에만 유효합니다.
        /// </summary>
        public string RefSchemaName
        {
            get;
            set;
        } = string.Empty;

        /// <summary>
        /// 가리키는 테이블을 반환합니다.
        /// 이 값은 REF 타입을 때에만 유효합니다.
        /// </summary>
        public string RefTableName
        {
            get;
            set;
        } = string.Empty;

        /// <summary>
        /// 가리키는 테이블에서 가리키는 행의 필드 이름을 반환합니다.
        /// </summary>
        public string RefFieldName
        {
            get;
            set;
        } = string.Empty;

        /// <summary>
        /// 가리키는 테이블 행에서 가져올 값.
        /// </summary>
        public string RefPickedFieldName
        {
            get;
            set;
        } = string.Empty;

        /// <summary>
        /// 가리키는 테이블 행에서 가져올 필드의 타입.
        /// </summary>
        public FieldType CachedRefPickedFieldType
        {
            get;
            set;
        } = FieldType.Nil;

        /// <summary>
        /// 이 필드의 그룹입니다.
        /// </summary>
        public HashSet<string> Groups
        {
            get;
            set;
        } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 필드 비활성화 태그들
        /// </summary>
        public HashSet<string> DisableTags
        {
            get;
            set;
        } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 이 필드의 컴포짓 키 그룹 이름 입니다.
        /// </summary>
        public HashSet<string> CompositKeys
        {
            get;
            set;
        } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 의존성 없는 타입 여부를 반환합니다.
        /// </summary>
        public bool IsSimpleType
        {
            get
            {
                switch(Type)
                {
                    case FieldType.Bool:
                    case FieldType.String:
                        return true;
                    default:
                        return IsIntegerType || IsRealType;
                }
            }
        }

        /// <summary>
        /// 부호 있는 정수 타입 여부 반환합니다.
        /// </summary>
        public bool IsIntegerType
        {
            get
            {
                switch (Type)
                {
                    case FieldType.Int8:
                    case FieldType.Int16:
                    case FieldType.Int32:
                    case FieldType.Int64:
                        return true;
                    default:
                        return false;
                }
            }
        }

        /// <summary>
        /// 실수형 타입 여부를 반환합니다.
        /// </summary>
        public bool IsRealType
        {
            get
            {
                switch (Type)
                {
                    case FieldType.Real32:
                    case FieldType.Real64:
                        return true;
                    default:
                        return false;
                }
            }
        }

        /// <summary>
        /// 컨테이너 타입 여부를 반환합니다.
        /// </summary>
        public bool IsContainerType
        {
            get
            {
                return Type == FieldType.List;
            }
        }

        /// <summary>
        /// 컨테이너 엘리먼트를 반환합니다. 
        /// 컨테이너 카테고리 타입인 경우에만 유효합니다.
        /// </summary>
        public IEnumerable<Field> Elements
        {
            get
            {
                return Enumerable.Repeat(ElementTemplate, IsContainerType ? ElementCount : 0);
            }
        }

        /// <summary>
        /// 컨테이너 엘리먼트 필드를 반환합니다.
        /// 이 값은 필드 타입이 컨테이너 타입일 경우에만 유효합니다.
        /// </summary>
        public Field ElementTemplate
        {
            get
            {
                return elementTemplate;
            }
            set
            {
                elementTemplate = value;
                if (elementTemplate != null)
                {
                    elementTemplate = value;
                    elementTemplate.NullDefaultValue = defaultValue;
                }
            }
        }

        /// <summary>
        /// 엘리먼트 개수를 반환합니다.
        /// 이 값은 필드 타입이 컨테이너 카테고리 타입일 경우에만 유효합니다.
        /// </summary>
        public int ElementCount
        {
            get;
            set;
        } = 0;

        /// <summary>
        /// 서로 호환 가능 여부를 반환합니다.
        /// </summary>
        /// <param name="other">비교할 필드</param>
        /// <returns>True, 호환 가능 시</returns>
        public bool IsCompatibleWith(Field other)
        {
            if (Type != other.Type)
            {
                return false;
            }

            //if (Type == FieldType.Ref && !DirectsSameRefenceTarget(other))
            //{
            //    return false;
            //}
            //else 
            if (Type == FieldType.Schema && !DirectsSameRefenceTarget(other))
            {
                return false;
            }
            else if (Type == FieldType.List)
            {
                if (ElementTemplate == null || other.ElementTemplate == null)
                {
                    return false;
                }

                if (ElementCount != other.ElementCount)
                {
                    return false;
                }

                if (!ElementTemplate.IsCompatibleWith(ElementTemplate))
                {
                    return false;
                }
            }

            if (AutoIncrement != other.AutoIncrement ||
                Unique != other.Unique ||
                NonSerialized != other.NonSerialized ||
                Nullable != other.Nullable)
            {
                return false;
            }

            if (!Groups.SequenceEqual(other.Groups) ||
                !DisableTags.SequenceEqual(other.DisableTags) ||
                !CompositKeys.SequenceEqual(other.CompositKeys))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 같은 참조 타겟 여부를 반환합니다.
        /// </summary>
        public bool DirectsSameRefenceTarget(Field other)
        {

            if (Type == FieldType.Schema && other.Type == FieldType.Schema)
            {
                return string.Equals(RefSchemaName, other.RefSchemaName, StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                return
                    string.Equals(RefTableName, other.RefTableName, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(RefFieldName, other.RefFieldName, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(RefPickedFieldName, other.RefPickedFieldName, StringComparison.OrdinalIgnoreCase);

            }
        }
    }
}
