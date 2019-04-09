using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;


namespace RandomSolutions
{
    public class ExcelConverter
    {
        public static T ConvertTo<T>(ExcelConvertMap map, string excelPath, string password = null)
        {
            return (T)ConvertTo(typeof(T), map, excelPath, password);
        }

        public static T ConvertTo<T>(ExcelConvertMap map, Stream excelStream, string password = null)
        {
            return (T)ConvertTo(typeof(T), map, excelStream, password);
        }

        public static object ConvertTo(Type type, ExcelConvertMap map, string excelPath, string password = null)
        {
            using (var stream = File.Open(excelPath, FileMode.Open, password != null ? FileAccess.ReadWrite : FileAccess.Read))
                return ConvertTo(type, map, stream, password);
        }

        public static object ConvertTo(Type type, ExcelConvertMap map, Stream excelStream, string password = null)
        {
            using (var excel = password != null ? new ExcelPackage(excelStream, password) : new ExcelPackage(excelStream))
            {
                excel.Compatibility.IsWorksheets1Based = true;
                return _getValue(excel, map, type);
            }
        }

        static object _getValue(ExcelPackage excel, ExcelConvertMap map, Type type, ExcelConvertContext parent = null)
        {
            var isEnumerable = map.Value == null && _isEnumerable(type);
            var itemType = !isEnumerable ? type : _getElementType(type);
            var itemPropInfos = itemType.GetProperties().Where(x => x.CanWrite).ToDictionary(x => x.Name, x => x);
            var count = isEnumerable && map.Break != null ? int.MaxValue : 1;
            ExcelConvertContext previous = null;

            var createCtx = new Func<ExcelConvertContext>(() => new ExcelConvertContext(excel)
            {
                Parent = parent,
                Previous = previous,
                Count = count,
                Sheet = parent?.Sheet ?? 1,
                Row = parent?.Row ?? 1,
                Col = parent?.Col ?? 1,
            });

            if (isEnumerable && map.Count != null)
                count = map.Count(createCtx());

            var items = new List<object>();

            for (var i = 0; i < count; i++)
            {
                var ctx = createCtx();
                ctx.Index = i;

                if (map.Sheet != null) ctx.Sheet = map.Sheet(ctx);
                if (map.Row != null) ctx.Row = map.Row(ctx);
                if (map.Col != null) ctx.Col = map.Col(ctx);

                if (!(ctx.Sheet > 0 && ctx.Row > 0 && ctx.Col > 0)
                    || ctx.Sheet > excel.Workbook.Worksheets.Count
                    || ctx.Col > excel.Workbook.Worksheets[ctx.Sheet].Dimension?.End.Column
                    || ctx.Row > excel.Workbook.Worksheets[ctx.Sheet].Dimension?.End.Row)
                    break;

                var ws = excel.Workbook.Worksheets[ctx.Sheet];
                var cell = ws.Cells[ctx.Row, ctx.Col];
                ctx.Value = cell.Value;

                if (map.Break?.Invoke(ctx) == true)
                    break;

                object item = null;

                if (map.Props == null)
                {
                    ctx.Value = map.Value != null ? map.Value(ctx) : _getCellValueSafe(cell, itemType);
                    item = ctx.Value;
                }
                else
                {
                    item = Activator.CreateInstance(itemType);

                    foreach (var propMap in map.Props)
                        if (itemPropInfos.ContainsKey(propMap.Name))
                        {
                            var pi = itemPropInfos[propMap.Name];
                            var val = _getValue(excel, propMap, pi.PropertyType, ctx);
                            pi.SetValue(item, val, null);
                        }
                }

                items.Add(item);
                previous = ctx;
            }

            return !isEnumerable ? (items.FirstOrDefault() ?? _getDefault(itemType))
                : type.IsArray ? _castTo(items, itemType, nameof(Enumerable.ToArray))
                : _castTo(items, itemType, nameof(Enumerable.ToList));
        }

        static bool _isEnumerable(Type type)
        {
            return type.IsArray
                || (type.IsInterface && type.IsGenericType && type.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                || (type.IsGenericType && type.GetInterfaces().Any(x => x.IsGenericType && x.GetGenericTypeDefinition() == typeof(IEnumerable<>)));
        }

        static bool _isEnum(Type type)
        {
            return type.IsEnum || (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>) && type.GetGenericArguments().FirstOrDefault()?.IsEnum == true);
        }

        static Type _getElementType(Type type)
        {
            return type.GetElementType() ??
                (type.IsInterface && type.IsGenericType ? type
                    : type.GetInterfaces().FirstOrDefault(x => x.IsGenericType && x.GetGenericTypeDefinition() == typeof(IEnumerable<>)))
                ?.GetGenericArguments().FirstOrDefault();
        }

        static object _getCellValueSafe(ExcelRange cell, Type type)
        {
            try
            {
                if (type.IsEnum)
                    return Enum.Parse(type, cell.Value?.ToString(), true);

                if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>) && type.GetGenericArguments().FirstOrDefault()?.IsEnum == true)
                    return Enum.Parse(type.GetGenericArguments().FirstOrDefault(), cell.Value?.ToString(), true);

                return typeof(ExcelRange).GetMethod(nameof(ExcelRangeBase.GetValue)).MakeGenericMethod(type).Invoke(cell, null);
            }
            catch (Exception ex)
            {
                return _getDefault(type);
            }
        }

        static object _getDefault(Type type) => new Func<object>(_getDefault<object>).Method.GetGenericMethodDefinition().MakeGenericMethod(type).Invoke(null, null);

        static T _getDefault<T>() => default(T);

        static object _castTo(object elements, Type elementType, string to)
        {
            var enumerable = typeof(Enumerable).GetMethod(nameof(Enumerable.Cast)).MakeGenericMethod(elementType).Invoke(elements, new[] { elements });
            return typeof(Enumerable).GetMethod(to).MakeGenericMethod(elementType).Invoke(enumerable, new[] { enumerable });
        }
    }
    
}
