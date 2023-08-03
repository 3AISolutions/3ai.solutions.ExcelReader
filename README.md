# 3ai.solutions.ExcelReader

## Sample read usage

```csharp
using var ms = new MemoryStream();
await context.HttpContext.Request.Body.CopyToAsync(ms);
ms.Position = 0;
var matrix = ms.ReadSingleExcelSheet().ToList();
```

## Sample reading fields

```csharp
int number = line.ElementAt(0).Value.ReadValue<int>();
string str = line.ElementAt(1).Value;
DateTime? date = line.ElementAt(2).Value.ReadNullableValue<DateTime>();
```
