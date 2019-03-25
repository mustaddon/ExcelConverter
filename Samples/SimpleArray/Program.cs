using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RandomSolutions;

namespace SimpleArray
{
    class Program
    {
        static void Main(string[] args)
        {
            var result = ExcelConverter.ConvertTo<SampleObject[]>(new ExcelConvertMap
            {
                Sheet = ctx => 1,
                Row = ctx => 3 + ctx.Index,
                Col = ctx => 2,
                Count = ctx => ctx.GetSheetEndRow(ctx.Sheet) - 2,
                Props = new[] {
                    new ExcelConvertMap {
                        Name = "Id",
                    },
                    new ExcelConvertMap {
                        Name = "Title",
                        Col = ctx => ctx.Parent.Col + 1,
                    },
                    new ExcelConvertMap {
                        Name = "Description",
                        Col = ctx => ctx.Parent.Col + 2,
                    },
                }
            }, @"sample.xlsx");

            foreach (var item in result)
                Console.WriteLine($"{item.Id}\t{item.Title}\t{item.Description}");

            Console.ReadKey();
        }
    }
}
