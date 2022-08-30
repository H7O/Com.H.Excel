﻿using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Dynamic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace Com.H.Excel
{
    internal class DataMapper
    {
        private readonly ConcurrentDictionary<Type, (string Name, PropertyInfo Info)[]> _typesProperties =
            new ConcurrentDictionary<Type, (string Name, PropertyInfo Info)[]>();


        public (string Name, PropertyInfo Info)[] GetCachedProperties(Type type)
        {
            if (type == null) throw new ArgumentNullException(nameof(type));
            return _typesProperties.ContainsKey(type) ?
                                _typesProperties[type]
                                : (_typesProperties[type] = GetPropertiesWithColumnName(type).ToArray());
        }
        public (string Name, PropertyInfo Info)[] GetCachedProperties(object obj)
        {
            if (obj == null) throw new ArgumentNullException(nameof(obj));
            if (obj.GetType() == typeof(ExpandoObject))
                return ((ExpandoObject)obj).GetProperties().ToArray();
            return GetCachedProperties(obj.GetType());
        }



        private static IEnumerable<(string Name, PropertyInfo Info)> GetPropertiesWithColumnName(Type type)
        {
            foreach (var p in type.GetProperties())
            {
                //ColumnAttribute c_name = (ColumnAttribute)p
                //    .GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault();
                // yield return (c_name == null ? p.Name : c_name.Name, p);
                yield return (p.Name, p);
            }

        }
        public IEnumerable<T> Map<T>(IEnumerable<object> source)
        {
            if (source == null) return null;
            return source.Select(x => this.Map<T>(x));
        }
        public object Map(object source, Type type)
            => type.GetMethod("Map").MakeGenericMethod(type)
            .Invoke(this, new object[] { source });

        public T Map<T>(object source)
        {
            if (source == null) return default;

            if (!typeof(IDictionary<string, object>).IsAssignableFrom(source.GetType()))
                return this.MapNormal<T>(source);

            IDictionary<string, object> srcProperties = source as IDictionary<string, object>;
            var dstProperties = this.GetCachedProperties(typeof(T));


            var joined = dstProperties.LeftJoin(
                srcProperties,
                dst => dst.Name.ToUpper(CultureInfo.InvariantCulture),
                src => src.Key.ToUpper(CultureInfo.InvariantCulture),
                (dst, src) => new { dst, src }
            );

            T destination = Activator.CreateInstance<T>();

            foreach (var item in joined.Where(x => x.src.Key != null))
            {
                try
                {
                    item.dst.Info.SetValue(destination,
                        Convert.ChangeType(item.src.Value,
                        item.dst.Info.PropertyType, CultureInfo.InvariantCulture)
                    );
                }
                catch { }
            }
            return destination;
        }

        private T MapNormal<T>(object source)
        {
            if (source == null) return default;

            var srcProperties = this.GetCachedProperties(source.GetType());
            var dstProperties = this.GetCachedProperties(typeof(T));
            // Console.WriteLine(dstProperties.Count());

            var joined = dstProperties.LeftJoin(
                srcProperties,
                dst => dst.Name.ToUpper(CultureInfo.InvariantCulture),
                src => src.Name.ToUpper(CultureInfo.InvariantCulture),
                (dst, src) => new { dst, src }
            ).Where(x => x.src.Info != null);

            // Console.WriteLine(joined.Count);

            T destination = Activator.CreateInstance<T>();

            foreach (var item in joined)
            {
                try
                {
                    // Console.WriteLine($"src: {item.src.Name} = {item.src.Info?.GetValue(source)}");
                    var val = item.src.Info.GetValue(source);

                    if (val is null) continue;
                    if (item.src.Info.PropertyType == item.dst.Info.PropertyType)
                        item.dst.Info.SetValue(destination, val);
                    else
                        item.dst.Info.SetValue(destination,
                            Convert.ChangeType(val,
                            item.dst.Info.PropertyType, CultureInfo.InvariantCulture)
                        );
                    // Console.WriteLine($"dst: {item.dst.Name} = {item.dst.Info?.GetValue(source)}");

                }
                catch // (Exception ex) 
                {
                    // Console.WriteLine("DataMapper: " + ex.Message);
                }
            }
            return destination;
        }


        public T Clone<T>(T source)
            => this.Map<T>(source);


        public void FillWith(
            object destination,
            object source,
            bool skipNull = false
            )
        {
            if (source == null || destination == null) return;

            var srcProperties = _typesProperties.ContainsKey(source.GetType()) ?
                                _typesProperties[source.GetType()]
                                : (_typesProperties[source.GetType()] =
                                    GetPropertiesWithColumnName(source.GetType()).ToArray());

            var dstProperties = _typesProperties.ContainsKey(destination.GetType()) ?
                                _typesProperties[destination.GetType()]
                                : (_typesProperties[destination.GetType()] =
                                GetPropertiesWithColumnName(destination.GetType()).ToArray());


            var joined = dstProperties.LeftJoin(
                srcProperties,
                dst => dst.Name.ToUpper(CultureInfo.InvariantCulture),
                src => src.Name.ToUpper(CultureInfo.InvariantCulture),
                (dst, src) => new { dst, src }
            );

            foreach (var item in joined.Where(x => (!skipNull || x.src.Info != null)))
            {
                try
                {
                    item.dst.Info.SetValue(destination,
                        item.src.Info == null ? null
                        :
                        Convert.ChangeType(item.src.Info.GetValue(source),
                        item.dst.Info.PropertyType, CultureInfo.InvariantCulture)
                    );
                }
                catch { }
            }
        }
    }
}