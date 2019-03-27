# Convert Excel to C# array with nested objects

![](/Samples/NestedArray/sample.png)

```C#
public class SampleObject
{
    public int Id { get; set; }
    public string Title { get; set; }
    public SampleUser User { get; set; }
    public SampleChild[] Childs { get; set; }
}

public class SampleUser
{
    public string Login { get; set; }
    public string Name { get; set; }
}

public class SampleChild
{
    public string Prop1 { get; set; }
    public string Prop2 { get; set; }
    public string[] Prop3 { get; set; }
}
```

```C#
var findRow = new Func<ExcelConvertContext, string, int>((ctx, title) 
    => ctx.FindRow(ctx.Sheet, ctx.Parent.Col, 
    (i, val) => val?.ToString() == title, ctx.Parent.Row, ctx.Parent.Row + 3));

var findChildCol = new Func<ExcelConvertContext, string, int>((ctx, title)
    => ctx.FindCol(ctx.Sheet, ctx.Parent.Parent.Row + 3, 
    (i, val) => val?.ToString() == title, ctx.Parent.Col, ctx.Parent.Col + 3));

var map = new ExcelConvertMap
{
    Sheet = ctx => 1,
    Row = ctx => 2,
    Col = ctx => 2 + 4 * ctx.Index,
    Break = ctx => ctx.Value == null,
    Props = new[] {
        new ExcelConvertMap {
            Name = "Id",
            Row = ctx => findRow(ctx, "id"),
            Col = ctx => ctx.Parent.Col + 1,
        },
        new ExcelConvertMap {
            Name = "Title",
            Row = ctx => findRow(ctx, "title"),
            Col = ctx => ctx.Parent.Col + 1,
        },
        new ExcelConvertMap {
            Name = "User",
            Row = ctx => findRow(ctx, "user"),
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
                    Col = ctx => findChildCol(ctx, "prop#1"),
                },
                new ExcelConvertMap {
                    Name = "Prop2",
                    Col = ctx => findChildCol(ctx, "prop#2"),
                },
                new ExcelConvertMap {
                    Name = "Prop3",
                    Col = ctx => findChildCol(ctx, "prop#3"),
		    Value = ctx => ctx.Value?.ToString().Split(','),
                },
            },
        },
    }
};

var result = ExcelConverter.ConvertTo<SampleObject[]>(map, @"sample.xlsx");
```
