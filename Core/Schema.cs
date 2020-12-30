using System;
using System.Collections.Generic;
using System.Linq;

namespace Hexx.Core
{
    public class Schema
    {
        HashSet<string> partials = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        List<Field> fields = new List<Field>();
        Dictionary<string, List<Field>> groups = new Dictionary<string, List<Field>>(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, List<Field>> compositeKeys = new Dictionary<string, List<Field>>(StringComparer.OrdinalIgnoreCase);

        public Schema()
        {
        }

        /// <summary>
        /// 스키마 복제본을 깊은 복사로 생성합니다.
        /// </summary>
        /// <param name="other"></param>
        public Schema(Schema other) :
            this(other.Name, from field in other.Fields
                             select new Field(field))
        {
            Description = other.Description;
            Embed = other.Embed;

            foreach(string name in other.Partials)
            {
                AddPartial(name);
            }
        }

        public Schema(string name, IEnumerable<Field> fields)
        {
            Name = name;
            foreach (Field field in fields)
            {
                AddField(field);
            }
        }

        /// <summary>
        /// 스키마 이름
        /// </summary>
        public string Name
        {
            get;
            set;
        } = string.Empty;

        /// <summary>
        /// 스키마에 대한 설명
        /// </summary>
        public string Description
        {
            get;
            set;
        } = string.Empty;

        /// <summary>
        /// 내장 스키마 여부
        /// </summary>
        public bool Embed
        {
            get;
            set;
        } = false;

        /// <summary>
        /// 스키마의 부분 테이블 이름을 반환합니다.
        /// </summary>
        public IEnumerable<string> Partials
        {
            get
            {
                return partials;
            }
        }

        /// <summary>
        /// 필드들을 반환합니다.
        /// </summary>
        public IEnumerable<Field> Fields
        {
            get
            {
                return fields;
            }
        }

        /// <summary>
        /// 필드 수
        /// </summary>
        public int FieldCount
        {
            get
            {
                return fields.Count;
            }
        }

        /// <summary>
        /// 스키마가 값 카테고리의 필드로만 이루어졌는지 여부를 반환합니다.
        /// 이 필드가 true를 반환할 때에만 테이블 화 할 수 있습니다.
        /// </summary>
        public bool IsFlat
        {
            get
            {
                return Fields.All(field => field.IsSimpleType);
            }
        }

        /// <summary>
        /// 특정 인덱스에 위치한 필드를 반환합니다.
        /// </summary>
        /// <param name="index">인덱스</param>
        /// <returns>필드. 유효한 인덱스가 아닐 경우 null이 반환 됩니다.</returns>
        public Field this[int index]
        {
            get
            {
                bool validRange = (index >= 0) && (index < fields.Count);
                return validRange ? fields[index] : null;
            }
        }

        /// <summary>
        /// 특정 이름의 필드를 반환합니다.
        /// </summary>
        /// <param name="fieldName">필드 이름</param>
        /// <returns>필드. 유효한 이름이 아닐 경우 null이 반환 됩니다.</returns>
        public Field this[string fieldName]
        {
            get
            {
                return GetField(fieldName);
            }
        }

        /// <summary>
        /// 호환되는 스키마 여부를 반환합니다.
        /// 서로 다른 테이블이라도 스키마만 호환 된다면, 서로의 테이블 행을 복사할 수 있습니다.
        /// </summary>
        /// <param name="fields"></param>
        /// <returns>스키마 호환 여부</returns>
        public bool IsCompatibleWith(IEnumerable<Field> fields)
        {
            if (this.fields.Count != fields.Count())
            {
                return false;
            }
            var otherEnumerator = fields.GetEnumerator();
            foreach (Field field in Fields)
            {
                if (!otherEnumerator.MoveNext())
                {
                    return false;
                }
                if (!field.IsCompatibleWith(otherEnumerator.Current))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// 스키마의 호환 여부를 반환합니다.
        /// 서로 다른 테이블이라도 스키마가 호환 된다면, 서로의 테이블 행을 복사할 수 있습니다.
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool IsCompatibleWith(Schema other)
        {
            return IsCompatibleWith(other.Fields);
        }

        /// <summary>
        /// 부분 테이블 이름을 추가합니다.
        /// </summary>
        /// <param name="name">이름</param>
        /// <returns>추가 성공 여부</returns>
        public bool AddPartial(string name)
        {
            return partials.Add(name);
        }

        /// <summary>
        /// 특정 이름의 필드 소유 여부를 반환합니다.
        /// </summary>
        /// <param name="name">필드 이름</param>
        /// <returns>필드 보유 여부</returns>
        public bool HasField(string name)
        {
            return GetFieldIndex(name) != -1;
        }

        /// <summary>
        /// 필드를 추가합니다.
        /// 같은 이름의 필드가 이미 있을 경우 아무것도 하지 않으며 false를 반환 합니다.
        /// </summary>
        /// <param name="newField">추가할 필드</param>
        /// <returns>추가 성공 여부</returns>
        public bool AddField(Field newField)
        {
            return AddFieldAt(fields.Count, newField);
        }

        /// <summary>
        /// 특정 인덱스 위치에 필드를 추가합니다.
        /// 같은 이름의 필드가 이미 있을 경우 아무것도 하지 않으며 false를 반환 합니다.
        /// </summary>
        /// <param name="index">인덱스</param>
        /// <param name="field">추가할 필드</param>
        /// <returns>필드 추가 성공 여부</returns>
        public bool AddFieldAt(int index, Field field)
        {
            bool validIdx = 0 <= index && index <= fields.Count;

            if (!validIdx || HasField(field.Name))
            {
                return false;
            }

            fields.Insert(index, field);

            foreach (string name in field.Groups)
            {
                List<Field> fields;
                if (!groups.TryGetValue(name, out fields))
                {
                    fields = new List<Field>();
                    groups.Add(name, fields);
                }
                fields.Add(field);
            }

            foreach (string name in field.CompositKeys)
            {
                List<Field> fields;
                if (!compositeKeys.TryGetValue(name, out fields))
                {
                    fields = new List<Field>();
                    compositeKeys.Add(name, fields);
                }
                fields.Add(field);
            }

            return true;
        }

        /// <summary>
        /// 필드의 컬럼 인덱스를 반환합니다.
        /// </summary>
        /// <param name="fieldName">반환할 필드 이름</param>
        /// <returns>필드 인덱스. 없는 필드 이름의 경우 -1이 반환 됩니다.</returns>
        public int GetFieldIndex(string fieldName)
        {
            return fields.FindIndex(field => field.Name.Equals(fieldName));
        }

        /// <summary>
        /// 특정 이름을 가진 필드를 반환합니다.
        /// 필드가 없다면 예외가 발생 합니다.
        /// </summary>
        /// <param name="fieldName">반환할 필드 이름</param>
        /// <returns>특정 이름의 필드. 필드가 없다면 null을 반환합니다.</returns>
        public Field GetField(string fieldName)
        {
            int fieldIndex = GetFieldIndex(fieldName);
            if (fieldIndex == -1)
            {
                return null;
            }

            return this[fieldIndex];
        }
    }
}
