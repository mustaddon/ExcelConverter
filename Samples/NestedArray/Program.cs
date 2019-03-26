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
            var result = ExcelConverter.ConvertTo<SampleObject[]>(new ExcelConvertMap
            {
                Sheet = ctx => 1,
                Row = ctx => 2,
                Col = ctx => 2 + 4 * ctx.Index,
                Break = ctx => ctx.Value == null,
                Props = new[] {
                    new ExcelConvertMap {
                        Name = "Id",
                        Row = ctx => FindRow(ctx, "id"),
                        Col = ctx => ctx.Parent.Col + 1,
                    },
                    new ExcelConvertMap {
                        Name = "Title",
                        Row = ctx => FindRow(ctx, "title"),
                        Col = ctx => ctx.Parent.Col + 1,
                    },
                    new ExcelConvertMap {
                        Name = "User",
                        Row = ctx => FindRow(ctx, "user"),
                        Props = new []{
                            new ExcelConvertMap {
                                Name = "Login",
                                Col = ctx => ctx.Parent.Col + 1,
                            },
                            new ExcelConvertMap {
                                Name = "Name",
                                Col = ctx => ctx.Parent.Col + 2,
                            },
                        },
                    },
                    new ExcelConvertMap {
                        Name = "Childs",
                        Row = ctx => ctx.Parent.Row + 4 + ctx.Index,
                        Break = ctx => ctx.Value == null,
                        Props = new []{
                            new ExcelConvertMap {
                                Name = "Prop1",
                                Col = ctx => FindChildCol(ctx, "prop#1"),
                            },
                            new ExcelConvertMap {
                                Name = "Prop2",
                                Col = ctx => FindChildCol(ctx, "prop#2"),
                            },
                            new ExcelConvertMap {
                                Name = "Prop3",
                                Col = ctx => FindChildCol(ctx, "prop#3"),
                                Value = ctx => ctx.Value?.ToString().Split(','),
                            },
                        },
                    },
                }
            }, @"sample.xlsx");

            foreach (var item in result)
            {
                Console.WriteLine($"{item.Id}\t{item.Title}\t{item.User.Login}\t{item.User.Name}");

                foreach(var child in item.Childs)
                    Console.WriteLine($"\t{child.Prop1}\t{child.Prop2}\t{string.Join(";",child.Prop3)}");

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
