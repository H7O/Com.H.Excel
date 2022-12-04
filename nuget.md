# Com.H.Excel
Kindly visit the project's github page for documentation [https://github.com/H7O/Com.H.Excel](https://github.com/H7O/Com.H.Excel)

## Examle 1
```csharp
using Com.H.Excel;
var list = new List<object>() {
	new { Name = "John", Age = 20 },
	new { Name = "Jane", Age = 21 },
	new { Name = "Jack", Age = 22 }
};
list.ToExcelFile("c:/temp/excel/excel01.xlsx");
```

## Example 2
```csharp
using Com.H.Excel;
var sheet1 = new List<object>() {
	new { Name = "John", Age = 20 },
	new { Name = "Jane", Age = 21 },
	new { Name = "Jack", Age = 22 }
};

var sheet2 = new List<object>() {
	new { Name = "Tom", Age = 20 },
	new { Name = "Helen", Age = 21 },
	new { Name = "Linda", Age = 22 },
};

var sheets = new Dictionary<string, IEnumerable<object>>() {
	{ "Sheet1", sheet1 },
	{ "Sheet2", sheet2 }
};

sheets.ToExcelFile("c:/temp/excel/excel02.xlsx");
```