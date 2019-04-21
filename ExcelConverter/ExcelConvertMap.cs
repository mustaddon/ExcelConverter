using System;


namespace RandomSolutions
{
    public class ExcelConvertMap
    {
        public string Name { get; set; }
        public Func<ExcelConvertContext, int> Sheet { get; set; }
        public Func<ExcelConvertContext, int> Row { get; set; }
        public Func<ExcelConvertContext, int> Col { get; set; }
        public Func<ExcelConvertContext, int> Count { get; set; }
        public Func<ExcelConvertContext, bool> Break { get; set; }
        public Func<ExcelConvertContext, object> Value { get; set; }
        public ExcelConvertMap[] Props { get; set; }
    }
}
