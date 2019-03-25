using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RandomSolutions;

namespace SimpleObject
{
    class Program
    {
        static void Main(string[] args)
        {
            var result = ExcelConverter.ConvertTo<SampleObject>(new ExcelConvertMap
            {
                Sheet = ctx => 1,
                Props = new[] {
                    new ExcelConvertMap {
                        Name = "Id",
                        Row = ctx => 2,
                        Col = ctx => 3,
                    },
                    new ExcelConvertMap {
                        Name = "Title",
                        Row = ctx => 3,
                        Col = ctx => 3,
                    },
                    new ExcelConvertMap {
                        Name = "Description",
                        Row = ctx => 4,
                        Col = ctx => 3,
                    },
                }
            }, @"sample.xlsx");

            Console.WriteLine(result.Id);
            Console.WriteLine(result.Title);
            Console.WriteLine(result.Description);
            Console.ReadKey();
        }
    }
}
