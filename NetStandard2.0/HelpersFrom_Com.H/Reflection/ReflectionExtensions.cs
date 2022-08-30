using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Com.H.Excel
{
    internal static class ReflectionExtensions
    {
        private static readonly DataMapper _mapper = new DataMapper();

        internal static (string Name, PropertyInfo Info)[] GetCachedProperties(this Type type)
            => _mapper.GetCachedProperties(type);

        internal static (string Name, PropertyInfo Info)[] GetCachedProperties(this object obj)
            => _mapper.GetCachedProperties(obj);

        internal static T Map<T>(this object source)
            => _mapper.Map<T>(source);

        internal static object Map(this object source, Type dstType)
            => _mapper.Map(source, dstType);

        internal static T Clone<T>(this T source)
            => _mapper.Clone<T>(source);

        internal static IEnumerable<T> Map<T>(this IEnumerable<object> source)
            => source==null?null:_mapper.Map<T>(source);

        internal static void FillWith(
            this object destination,
            object source,
            bool skipNull = false
            )
            => _mapper.FillWith(destination, source, skipNull);

        /// <summary>
        /// Rrturns values of IDictionary after filtering them based on an IEnumerable of keys.
        /// The filter keys don't have to be of the same type as the IDictionary keys.
        /// They only need to be mappable to IDictionary keys type (i.e. can be conerted to IDicionary keys type)
        /// </summary>
        /// <typeparam name="TKey"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <typeparam name="TOKey"></typeparam>
        /// <param name="dictionary"></param>
        /// <param name="oFilter"></param>
        /// <returns></returns>
        internal static IEnumerable<TValue> OrdinallyMappedFilteredValues<TKey, TValue, TOKey>(
            IDictionary<TKey, TValue> dictionary, IEnumerable<TOKey> oFilter) 
        =>
            oFilter is null?dictionary.Values.AsEnumerable()
            :oFilter.Where(x=>x != null).Join(dictionary, o => o.Map<TKey>(), d => d.Key, (o, d) => d.Value);



        internal static IEnumerable<(string Name, PropertyInfo Info)> GetProperties(this ExpandoObject expando)
        {
            if (expando == null) throw new ArgumentNullException(nameof(expando));
            foreach (var p in expando)
            {
                yield return (p.Key, new DynamicPropertyInfo(p.Key, p.Value?.GetType() ?? typeof(string)));
            }
        }


        internal static object GetDefault(this Type type)
            => ((Func<object>)GetDefault<object>)?.Method?.GetGenericMethodDefinition()?
            .MakeGenericMethod(type)?.Invoke(null, null);

        private static T GetDefault<T>()
            => default;


        internal static object ConvertTo(this object obj, Type type)
        {
            Type dstType = Nullable.GetUnderlyingType(type) ?? type;
            return (obj == null || DBNull.Value.Equals(obj)) ?
               type.GetDefault() : Convert.ChangeType(obj, dstType);
        }

        internal static bool IsDefault<T>(this T value) where T : struct
            => value.Equals(default(T));
    }
}