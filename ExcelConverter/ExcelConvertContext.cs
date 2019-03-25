using OfficeOpenXml;
using System;
using System.Linq;


namespace RandomSolutions
{

    public class ExcelConvertContext
    {
        private ExcelPackage _excel;

        internal ExcelConvertContext(ExcelPackage excel)
        {
            _excel = excel;
        }

        public ExcelConvertContext Parent { get; internal set; }
        public ExcelConvertContext Previous { get; internal set; }
        public int Count { get; internal set; }
        public int Index { get; internal set; } = -1;
        public int Sheet { get; internal set; } = -1;
        public int Row { get; internal set; } = -1;
        public int Col { get; internal set; } = -1;
        public object Value { get; internal set; }


        public int GetSheetStartRow(int sheet)
            => _excel.Workbook.Worksheets[sheet].Dimension?.Start.Row ?? 1;

        public int GetSheetStartCol(int sheet)
            => _excel.Workbook.Worksheets[sheet].Dimension?.Start.Column ?? 1;

        public int GetSheetEndRow(int sheet)
            => _excel.Workbook.Worksheets[sheet].Dimension?.End.Row ?? 1;

        public int GetSheetEndCol(int sheet)
            => _excel.Workbook.Worksheets[sheet].Dimension?.End.Column ?? 1;

        public object GetValue(int sheet, int row, int col)
            => _excel.Workbook.Worksheets[sheet].Cells[row, col].Value;

        public T GetValue<T>(int sheet, int row, int col)
            => _excel.Workbook.Worksheets[sheet].Cells[row, col].GetValue<T>();

        public T GetValue<T>()
            => _excel.Workbook.Worksheets[Sheet].Cells[Row, Col].GetValue<T>();


        public int FindSheet(string name)
            => _excel.Workbook.Worksheets[name]?.Index ?? -1;

        public int FindSheet(Func<int, string, bool> test)
            => _excel.Workbook.Worksheets.FirstOrDefault(x => test(x.Index, x.Name))?.Index ?? -1;

        public int FindRow(int sheet, int col, Func<int, object, bool> test,
            int? rowStart = null, int? rowEnd = null)
            => FindRow(sheet, (r, c, v) => test(r, v), rowStart, rowEnd, col, col);

        public int FindRow(int sheet, Func<int, int, object, bool> test,
            int? rowStart = null, int? rowEnd = null,
            int? colStart = null, int? colEnd = null)
        {
            var ws = _excel.Workbook.Worksheets[sheet];
            var sRow = rowStart ?? GetSheetStartRow(sheet);
            var eRow = rowEnd ?? GetSheetEndRow(sheet);
            var sCol = colStart ?? GetSheetStartCol(sheet);
            var eCol = colEnd ?? GetSheetEndCol(sheet);

            for (var col = sCol; col <= eCol; col++)
                for (var row = sRow; row <= eRow; row++)
                    if (test(row, col, ws.Cells[row, col]?.Value))
                        return row;

            return -1;
        }

        public int FindCol(int sheet, int row, Func<int, object, bool> test,
            int? colStart = null, int? colEnd = null)
            => FindCol(sheet, (r, c, v) => test(c, v), row, row, colStart, colEnd);

        public int FindCol(int sheet, Func<int, int, object, bool> test,
            int? rowStart = null, int? rowEnd = null,
            int? colStart = null, int? colEnd = null)
        {
            var ws = _excel.Workbook.Worksheets[sheet];
            var sRow = rowStart ?? GetSheetStartRow(sheet);
            var eRow = rowEnd ?? GetSheetEndRow(sheet);
            var sCol = colStart ?? GetSheetStartCol(sheet);
            var eCol = colEnd ?? GetSheetEndCol(sheet);

            for (var row = sRow; row <= eRow; row++)
                for (var col = sCol; col <= eCol; col++)
                    if (test(row, col, ws.Cells[row, col]?.Value))
                        return col;

            return -1;
        }


    }
}
