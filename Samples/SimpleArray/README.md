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
```
