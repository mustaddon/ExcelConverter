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
var map = new ExcelConvertMap
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
};

var result = ExcelConverter.ConvertTo<SampleObject>(map, @"sample.xlsx");
```
