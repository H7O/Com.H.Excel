# Com.H.Excel
Allows for easy creation and consumption of xlsx files. 

## Installation
The easiest way to install Com.H.Excel is via NuGet GUI or CLI.

### NuGet CLI
In terminal, under your project folder (not solution folder) run the following:
```
dotnet add package Com.H.Excel
```

## How to use

### Examle 1
Writing a single sheet excel

```csharp
using Com.H.Excel;
var list = new List<object>() {
	new { Name = "John", Age = 20 },
	new { Name = "Jane", Age = 21 },
	new { Name = "Jack", Age = 22 }
};
list.ToExcelFile("c:/temp/excel/excel01.xlsx");
```

### Example 2
Writing multi-sheeet excel.

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

### Example 3
Getting a stream reader to a generated excel temp file that gets automatically deleted once the reader is closed.

```csharp
using Com.H.Excel;
var list = new List<object>() {
	new { Name = "John", Age = 20 },
	new { Name = "Jane", Age = 21 },
	new { Name = "Jack", Age = 22 }
};
var stream = list.ToExcelReader();
```






