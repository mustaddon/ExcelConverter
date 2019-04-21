# Convert Excel to simple C# object

![](/Samples/SimpleObject/sample.png)

```C#
public class SampleObject
{
    public int Id { get; set; }
    public string Title { get; set; }
    public string Description { get; set; }
}
```

```C#
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
```
