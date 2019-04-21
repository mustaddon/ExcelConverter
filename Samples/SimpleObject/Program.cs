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
            var result = ExcelConverter.ConvertTo<SampleObject>(a => a
                .Sheet(x => 1)
                .Prop("Id", b => b
                    .Row(x => 2)
                    .Col(x => 3))
                .Prop("Title", b => b
                    .Row(x => 3)
                    .Col(x => 3))
                .Prop("Description", b => b
                    .Row(x => 4)
                    .Col(x => 3))
            , @"sample.xlsx");


            Console.WriteLine(result.Id);
            Console.WriteLine(result.Title);
            Console.WriteLine(result.Description);
            Console.ReadKey();
        }
    }
}
