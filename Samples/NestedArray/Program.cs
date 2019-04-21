using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RandomSolutions;

namespace NestedArray
{
    class Program
    {
        static void Main(string[] args)
        {
            var result = ExcelConverter.ConvertTo<SampleObject[]>(a => a
                .Sheet(x => 1)
                .Row(x => 2)
                .Col(x => 2 + 4 * x.Index)
                .Break(x => x.Value == null)
                .Prop("Id", b => b
                    .Row(x => FindRow(x, "id"))
                    .Col(x => x.Parent.Col + 1))
                .Prop("Title", b => b
                    .Row(x => FindRow(x, "title"))
                    .Col(x => x.Parent.Col + 1))
                .Prop("User", b => b
                    .Row(x => FindRow(x, "user"))
                    .Prop("Login", c => c
                        .Col(x => x.Parent.Col + 1))
                    .Prop("Name", c => c
                        .Col(x => x.Parent.Col + 2)))
                .Prop("Childs", b => b
                    .Row(x => x.Parent.Row + 4 + x.Index)
                    .Break(x => x.Value == null)
                    .Prop("Prop1", c => c
                        .Col(x => FindChildCol(x, "prop#1")))
                    .Prop("Prop2", c => c
                        .Col(x => FindChildCol(x, "prop#2")))
                    .Prop("Prop3", c => c
                        .Col(x => FindChildCol(x, "prop#3"))
                        .Value(x => x.Value?.ToString().Split(','))))
            , @"sample.xlsx");

            foreach (var item in result)
            {
                Console.WriteLine($"{item.Id}\t{item.Title}\t{item.User.Login}\t{item.User.Name}");

                foreach (var child in item.Childs)
                    Console.WriteLine($"\t{child.Prop1}\t{child.Prop2}\t{string.Join(";", child.Prop3)}");

                Console.WriteLine();
            }

            Console.ReadKey();
        }

        static int FindRow(ExcelConvertContext ctx, string title)
        {
            return ctx.FindRow(ctx.Sheet, ctx.Parent.Col,
                (i, v) => v?.ToString() == title,
                ctx.Parent.Row, ctx.Parent.Row + 3);
        }

        static int FindChildCol(ExcelConvertContext ctx, string title)
        {
            return ctx.FindCol(ctx.Sheet, ctx.Parent.Parent.Row + 3,
                (i, v) => v?.ToString() == title,
                ctx.Parent.Col, ctx.Parent.Col + 3);
        }
    }
}
