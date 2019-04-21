using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RandomSolutions
{
    public class ExcelConvertMapBuilder
    {
        ExcelConvertMap _map = new ExcelConvertMap();
        public Dictionary<string, ExcelConvertMapBuilder> _props = new Dictionary<string, ExcelConvertMapBuilder>();
        
        public ExcelConvertMapBuilder Count(Func<ExcelConvertContext, int> fn)
        {
            _map.Count = fn;
            return this;
        }

        public ExcelConvertMapBuilder Sheet(Func<ExcelConvertContext, int> fn)
        {
            _map.Sheet = fn;
            return this;
        }

        public ExcelConvertMapBuilder Row(Func<ExcelConvertContext, int> fn)
        {
            _map.Row = fn;
            return this;
        }

        public ExcelConvertMapBuilder Col(Func<ExcelConvertContext, int> fn)
        {
            _map.Col = fn;
            return this;
        }

        public ExcelConvertMapBuilder Value(Func<ExcelConvertContext, object> fn)
        {
            _map.Value = fn;
            return this;
        }

        public ExcelConvertMapBuilder Break(Func<ExcelConvertContext, bool> fn)
        {
            _map.Break = fn;
            return this;
        }

        public ExcelConvertMapBuilder Prop(string propName, Action<ExcelConvertMapBuilder> propMapBuilder = null)
        {
            var builder = new ExcelConvertMapBuilder();
            propMapBuilder?.Invoke(builder);

            if (!_props.ContainsKey(propName))
                _props.Add(propName, builder);
            else
                _props[propName] = builder;

            return this;
        }

        public ExcelConvertMap Build()
        {
            _map.Props = !_props.Any() ? null
                : _props.Select(x => {
                    var propMap = x.Value.Build();
                    propMap.Name = x.Key;
                    return propMap;
                }).ToArray();

            return _map;
        }
    }
}
