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
            var result = ExcelConverter.ConvertTo<SampleObject[]>(a => a
                .Sheet(x => 1)
                .Row(x => 3 + x.Index)
                .Col(x => 2)
                .Count(x => x.GetSheetEndRow(x.Sheet) - 2)
                .Prop("Id")
                .Prop("Title", b => b
                    .Col(x => x.Parent.Col + 1))
                .Prop("Description", b => b
                    .Col(x => x.Parent.Col + 2))
            , @"sample.xlsx");

            foreach (var item in result)
                Console.WriteLine($"{item.Id}\t{item.Title}\t{item.Description}");

            Console.ReadKey();
        }
    }
}
