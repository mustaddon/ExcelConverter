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

 var result = ExcelConverter.ConvertTo<SampleObject[]>(a => a
    .Sheet(x => 1)
    .Row(x => 2)
    .Col(x => 2 + 4 * x.Index)
    .Break(x => x.Value == null)
    .Prop("Id", b => b
        .Row(x => findRow(x, "id"))
        .Col(x => x.Parent.Col + 1))
    .Prop("Title", b => b
        .Row(x => findRow(x, "title"))
        .Col(x => x.Parent.Col + 1))
    .Prop("User", b => b
        .Row(x => findRow(x, "user"))
        .Prop("Login", c => c
            .Col(x => x.Parent.Col + 1))
        .Prop("Name", c => c
            .Col(x => x.Parent.Col + 2)))
    .Prop("Childs", b => b
        .Row(x => x.Parent.Row + 4 + x.Index)
        .Break(x => x.Value == null)
        .Prop("Prop1", c => c
            .Col(x => findChildCol(x, "prop#1")))
        .Prop("Prop2", c => c
            .Col(x => findChildCol(x, "prop#2")))
        .Prop("Prop3", c => c
            .Col(x => findChildCol(x, "prop#3"))
            .Value(x => x.Value?.ToString().Split(',')))
    )
, @"sample.xlsx");
```
