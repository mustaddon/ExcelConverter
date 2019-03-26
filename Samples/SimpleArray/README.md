# Convert Excel to simple C# array

![](/Samples/SimpleArray/sample.png)

```C#
public class SampleObject
{
    public int Id { get; set; }
    public string Title { get; set; }
    public string Description { get; set; }
}
```

```C#
var map = new ExcelConvertMap
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
};

var result = ExcelConverter.ConvertTo<SampleObject[]>(map, @"sample.xlsx");
```
