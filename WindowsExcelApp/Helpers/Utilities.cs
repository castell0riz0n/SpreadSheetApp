using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace WindowsExcelApp.Helpers
{
    public static class Utilities
    {

        public static bool IsJson(this object objectModel)
        {
            bool result = false;
            try
            {
                if (objectModel != null && ((objectModel.ToString().StartsWith("[") && objectModel.ToString().EndsWith("]")) || (objectModel.ToString().StartsWith("{") && objectModel.ToString().EndsWith("}"))))
                {
                    result = true;
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public static string ToJson(this object objectModel)
        {
            try
            {
                if (objectModel == null)
                    return "null";

                var jsonSetting = new JsonSerializerSettings
                {
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                };

                return Newtonsoft.Json.JsonConvert.SerializeObject(objectModel, Formatting.None, jsonSetting);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public static string StringJoin(this string[] values)
        {
            try
            {
                if (values.IsNullOrEmpty())
                    return "";

                return string.Join(",", values);
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        public static string StringJoin(this List<string> values)
        {
            return StringJoin(values.ToArray());
        }

        public static string ToJson(this object objectModel, Type type)
        {
            try
            {
                var jsonSetting = new JsonSerializerSettings
                {
                    //ReferenceLoopHandling = Newtonsoft.Json.ReferenceLoopHandling.Ignore,
                };

                return Newtonsoft.Json.JsonConvert.SerializeObject(objectModel, type, Formatting.None, jsonSetting);
            }
            catch (Exception ex)
            {


                return ex.Message;
            }
        }

        public static object ToObjectFromJson(this string objectModel, Type type)
        {
            try
            {
                if (string.IsNullOrEmpty(objectModel))
                    return null;

                var settings = GetJsonSerializerSettings();

                return Newtonsoft.Json.JsonConvert.DeserializeObject(objectModel, type, settings);
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private static JsonSerializerSettings GetJsonSerializerSettings()
        {
            return new JsonSerializerSettings()
            {
                NullValueHandling = NullValueHandling.Ignore
            };
        }

        public static T ToObjectFromJson<T>(this string objectModel)
        {
            try
            {
                if (string.IsNullOrEmpty(objectModel))
                    return default(T);

                var settings = GetJsonSerializerSettings();

                return (T)Newtonsoft.Json.JsonConvert.DeserializeObject<T>(objectModel, settings);
            }
            catch (Exception ex)
            {
                return default(T);
            }
        }

        public static Guid ToGuid(this object objectModel)
        {
            try
            {
                if (objectModel == null)
                    return default(Guid);

                return Guid.Parse(objectModel.ToString());
            }
            catch (Exception ex)
            {
                return default(Guid);
            }
        }

        public static bool IsNullOrEmpty<T>(this IEnumerable<T> enumerable)
        {
            if (enumerable == null)
            {
                return true;
            }

            var collection = enumerable as ICollection<T>;

            if (collection != null)
            {
                return collection.Count < 1;
            }
            return !enumerable.Any();
        }

        public static bool HasCountItem<T>(this IEnumerable<T> enumerable, int countItem)
        {
            if (IsNullOrEmpty<T>(enumerable))
                return false;

            return enumerable.Count() == countItem;
        }

        public static bool IsNullOrEmpty(this Guid guid)
        {
            if (guid == null)
            {
                return true;
            }

            return guid == Guid.Empty;
        }

        public static bool IsNullOrEmpty(this Guid? guid)
        {
            if (!guid.HasValue)
            {
                return true;
            }

            return guid == Guid.Empty;
        }

        public static bool IsNullOrEmpty(this DateTime dateTime)
        {
            if (dateTime == null)
                return true;

            if (dateTime.Year == 1)
                return true;

            if (dateTime <= DateTime.MinValue)
                return true;

            return false;
        }

        public static bool IsNullOrEmpty(this DateTime? dateTime)
        {
            if (!dateTime.HasValue)
                return true;

            if (dateTime.Value.Year == 1)
                return true;

            if (dateTime.Value <= DateTime.MinValue)
                return true;

            return false;
        }

        public static bool IsNullOrEmpty(this string value)
        {
            return string.IsNullOrEmpty(value);
        }

        public static T ConvertToEnumFromString<T>(this string value)
        where T : struct, IConvertible
        {
            var sourceType = value.GetType();

            if (!typeof(T).IsEnum)
                throw new ArgumentException("Destination type is not enum");

            return ConvertToEnumWithStringValue<T>(value);
        }

        private static T ConvertToEnumWithStringValue<T>(string value) where T : struct, IConvertible
        {
            try
            {
                return (T)Enum.Parse(typeof(T), value.ToString());
            }
            catch (Exception)
            {
                foreach (var val in Enum.GetValues(typeof(T)))
                {
                    if (val.ToString().ToLower() == value.NullToEmpty().ToLower())
                        return (T)val;
                }
            }

            return default(T);
        }

        public static T ConvertToEnum<T>(this object value)
        where T : struct, IConvertible
        {
            var sourceType = value.GetType();

            if (!sourceType.IsEnum)
                throw new ArgumentException("Source type is not enum");

            if (!typeof(T).IsEnum)
                throw new ArgumentException("Destination type is not enum");

            return (T)Enum.Parse(typeof(T), value.ToString());
        }

        public static T ConvertTo<T>(this object value)
        {
            return (T)value;
        }

        public static T SetIdProperty<T>(T entity, object value)
        {
            Type entityType = entity.GetType();

            string propertyName = $"{entityType.Name}Id";

            return SetProperty<T>(entity, value, propertyName);
        }

        public static T GetIdProperty<T>(object entity)
        {
            Type entityType = entity.GetType();

            string propertyName = $"{entityType.Name}Id";

            return GetProperty<T>(entity, propertyName);
        }

        public static object GetIdProperty(object entity)
        {
            Type entityType = entity.GetType();

            string propertyName = $"{entityType.Name}Id";

            return GetProperty(entity, propertyName);
        }

        public static T SetProperty<T>(this T entity, object value, string propertyName)
        {
            Type entityType = entity.GetType();

            PropertyInfo prop = entityType.GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);

            if (null != prop && prop.CanWrite)
            {
                prop.SetValue(entity, value, null);
            }

            return entity;
        }

        public static object GetProperty(this object obj, string propertyName)
        {
            Type myType = obj.GetType();

            IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());

            foreach (PropertyInfo prop in props)
            {
                if (prop.Name == propertyName)
                {
                    return prop.GetValue(obj, null);
                }
            }

            return null;
        }

        public static T GetProperty<T>(this object obj, string propertyName)
        {
            Type myType = obj.GetType();

            IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());

            foreach (PropertyInfo prop in props)
            {
                if (prop.Name == propertyName)
                {
                    return (T)prop.GetValue(obj, null);
                }
            }

            return default(T);
        }

        public static bool HasProperty(this object obj, string propertyName)
        {
            Type myType = obj.GetType();

            return HasProperty(myType, propertyName);

        }

        public static bool HasProperty(Type objectType, string propertyName)
        {
            Type myType = objectType;

            IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());

            return props.Any(p => p.Name == propertyName);

        }

        public static bool HasProperty<T>(string propertyName)
        {
            Type myType = typeof(T);

            IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());

            return props.Any(p => p.Name == propertyName);

        }

        public static Dictionary<int, string> EnumToDictionary<T>()
        {
            if (!typeof(T).IsEnum)
            {
                return null;
            }

            var dict = Enum.GetValues(typeof(T))
               .Cast<T>()
               .ToDictionary(t => (int)(dynamic)t, t => GetEnumDescription<T>(t.ToString()));

            return dict;
        }

        public static Dictionary<object, string> EnumToDictionary(string enumName, bool numKey = true)
        {
            foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                var type = assembly.GetType("GeneralWebsite.Infostructure." + enumName);
                if (type != null)
                {
                    if (type.IsEnum)
                    {
                        var dict = Enum.GetValues(type)
                           .EnumArrayToDictionary(type, numKey);

                        return dict;
                    }
                    return null;
                }
            }
            return null;
        }

        public static Dictionary<object, string> EnumArrayToDictionary(this System.Collections.IEnumerable source, Type enumType, bool numKey = true)
        {
            var result = new Dictionary<object, string>();

            foreach (var item in source)
            {
                var enumVariable = Enum.Parse(enumType, item.ToString());

                result.Add(numKey ? enumVariable.GetHashCode().ToString() : enumVariable.ToString(), GetEnumDescription(enumVariable.ToString(), enumType));
            }

            return result;
        }

        public static string GetEnumDescription(this Enum value)
        {
            Type type = value.GetType();

            return GetEnumDescription(value.ToString(), type);
        }

        public static string GetEnumDescription<T>(string value)
        {
            Type type = typeof(T);

            return GetEnumDescription(value, type);
        }

        private static string GetEnumDescription(string value, Type type)
        {
            try
            {
                var name = Enum.GetNames(type).Where(f => f.Equals(value, StringComparison.CurrentCultureIgnoreCase)).Select(d => d).FirstOrDefault();

                if (!name.IsNullOrEmpty())
                {
                    var field = type.GetField(name);

                    var customAttribute = field.GetCustomAttributes(typeof(DescriptionAttribute), false);

                    if (!customAttribute.IsNullOrEmpty())
                        return customAttribute.Length > 0 ? ((DescriptionAttribute)customAttribute[0]).Description : name;

                    return name;
                }
            }
            catch (Exception)
            {
            }

            return string.Empty;
        }

        public static byte GetEnumValueFromDescription<T>(string description)
        {
            var type = typeof(T);
            if (!type.IsEnum) throw new InvalidOperationException();
            foreach (var field in type.GetFields())
            {
                var attribute = Attribute.GetCustomAttribute(field,
                    typeof(DescriptionAttribute)) as DescriptionAttribute;
                if (attribute != null)
                {
                    if (attribute.Description == description)
                        return (byte)field.GetValue(null);
                }
                else
                {
                    if (field.Name == description)
                        return (byte)field.GetValue(null);
                }
            }
            throw new ArgumentException("Not found.", "description");
        }

        public static T GetEnumFromString<T>(string value)
            where T : struct, IComparable
        {
            try
            {
                if (!value.IsNullOrEmpty())
                {
                    Type type = typeof(T);

                    if (type.IsEnum)
                    {
                        T enumTemp;
                        //enumTemp = (T)Enum.Parse(type, value);

                        if (!Enum.TryParse<T>(value, out enumTemp))
                        {
                            if (enumTemp.GetHashCode() == 0)
                                enumTemp = (T)Enum.Parse(type, value.UppercaseFirst());
                        }

                        return enumTemp;
                    }
                }
            }
            catch (Exception ex)
            {
            }

            return default(T);
        }

        public static string UppercaseFirst(this string valus)
        {
            if (string.IsNullOrEmpty(valus))
            {
                return string.Empty;
            }

            char[] a = valus.ToLower().ToCharArray();

            a[0] = char.ToUpper(a[0]);

            return new string(a);
        }

        public static string ToUpperFirstChar(this string input)
        {
            switch (input)
            {
                case null: return input;
                case "": return input;
                default: return input.First().ToString().ToUpper() + input.Substring(1);
            }
        }

        public static string NullToEmpty(this string value)
        {
            return value ?? "";
        }

        public static byte GetByteCode(this Enum enumParameter)
        {
            return (byte)enumParameter.GetHashCode();
        }

        public static int GetInt32Code(this Enum enumParameter)
        {
            return (int)enumParameter.GetHashCode();
        }

        public static bool AreSameType<TType>(this object obj)
        {
            if (obj.GetType() == typeof(TType))
                return true;

            return false;
        }

        public static T CloneObjectReflection<T>(this object objSource)
        {
            if (objSource == null)
                return default(T);

            return (T)CloneObjectReflection(objSource);
        }

        public static object CloneObjectReflection(object obj)
        {
            try
            {
                if (obj == null)
                    return null;

                Type type = obj.GetType();

                if (type.IsValueType || type == typeof(string))
                {
                    return obj;
                }
                else if (type.IsArray)
                {
                    var typeNameString = $"{type.FullName.Replace("[]", string.Empty)}, {type.Assembly.FullName}";

                    Type elementType = Type.GetType(typeNameString);

                    if (elementType == null)
                        return null;

                    var array = obj as Array;

                    Array copied = Array.CreateInstance(elementType, array.Length);

                    for (int i = 0; i < array.Length; i++)
                    {
                        copied.SetValue(CloneObjectReflection(array.GetValue(i)), i);
                    }

                    return Convert.ChangeType(copied, obj.GetType());
                }
                else if (type.IsClass)
                {
                    object toret = Activator.CreateInstance(obj.GetType());

                    FieldInfo[] fields = type.GetFields();

                    //FieldInfo[] fields = type.GetFields(BindingFlags.Public |
                    //            BindingFlags.NonPublic | BindingFlags.Instance);
                    foreach (FieldInfo field in fields)
                    {
                        object fieldValue = field.GetValue(obj);

                        if (fieldValue == null)
                            continue;

                        field.SetValue(toret, CloneObjectReflection(fieldValue));
                    }

                    PropertyInfo[] properties = type.GetProperties();

                    foreach (PropertyInfo property in properties)
                    {
                        object propertyValue = property.GetValue(obj);

                        if (propertyValue == null)
                            continue;

                        property.SetValue(toret, CloneObjectReflection(propertyValue));
                    }

                    return toret;
                }
                else
                    return null;

            }
            catch (Exception)
            {
                return null;
            }
        }

        public static T CloneObjectJson<T>(this object objSource)
        {
            var json = objSource.ToJson();

            return json.ToObjectFromJson<T>();
        }

        public static object CloneObjectJson(this object objSource)
        {
            try
            {
                var json = objSource.ToJson();

                var objectType = objSource.GetType();

                var result = json.ToObjectFromJson(objectType);

                return result;
            }
            catch (Exception ex)
            {

                return null;
            }

        }

        public static int ToInt32(this object value)
        {
            if (value == null)
                return 0;

            int result = 0;

            int.TryParse(value.ToString(), out result);

            return result;
        }

        public static long ToInt64(this object value)
        {
            long result = 0;

            long.TryParse(value.ToString(), out result);

            return result;
        }

        public static decimal ToDecimal(this object value)
        {
            decimal result = 0;

            decimal.TryParse(value.ToString(), out result);

            return result;
        }

        public static byte ToByte(this object value)
        {
            byte result = 0;

            byte.TryParse(value.ToString(), out result);

            return result;
        }

        public static bool ToBoolean(this object value)
        {
            if (value == null)
                return false;

            bool result = false;

            bool.TryParse(value.ToString(), out result);

            return result;
        }

        public static DateTime ToDateTime(this object value)
        {
            if (value == null)
                return DateTime.MaxValue;

            if (string.IsNullOrEmpty(value.ToString()))
                return DateTime.MaxValue;

            DateTime.TryParse(value.ToString(), out DateTime result);

            return result;
        }

        public static void RenameKey<TKey, TValue>(this IDictionary<TKey, TValue> dic,
            TKey fromKey, TKey toKey)
        {
            TValue value = dic[fromKey];
            dic.Remove(fromKey);
            dic[toKey] = value;
        }

        public static IEnumerable<TSource> DistinctBy<TSource, TKey>
            (this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        {
            HashSet<TKey> seenKeys = new HashSet<TKey>();
            foreach (TSource element in source)
            {
                if (seenKeys.Add(keySelector(element)))
                {
                    yield return element;
                }
            }
        }

        public static string ReplaceChar(this object objectModel)
        {
            try
            {
                if (objectModel == null)
                    return "";

                return objectModel.ToString().Replace("{", "").Replace("}", "").Replace(" ", "").Trim();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public static T ToEnum<T>(this string value, T defaultValue)
        {
            if (string.IsNullOrEmpty(value))
            {
                return defaultValue;
            }

            //T result;
            //return Enum.TryParse<T>(value, true, out result) ? result : defaultValue;
            return (T)Enum.Parse(typeof(T), value, true);
        }


        public static byte[] ObjectToByteArray(this object value)
        {
            byte[] result = null;
            if (value == null)
                return null;

            var binFormatter = new BinaryFormatter();

            using (var memoryStream = new MemoryStream())
            {
                binFormatter.Serialize(memoryStream, value);

                result = memoryStream.ToArray();
            }
            return result;
        }


        public static bool WordIsInPersianOrArabic(string input)
        {
            Regex regex = new Regex("[\u0600-\u06ff]|[\u0750-\u077f]|[\ufb50-\ufc3f]|[\ufe70-\ufefc]");
            return regex.IsMatch(input);
        }
    }
}