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

### Writing Excel

#### Sample 1
Writing a single sheet excel

```csharp
using Com.H.Excel;
// Note: you can use a specific class instead of anonymous object. 
// e.g. new List<Person>()
var list = new List<object>() {
	new { Name = "John", Age = 20 },
	new { Name = "Jane", Age = 21 },
	new { Name = "Jack", Age = 22 }
};
list.ToExcelFile("c:/temp/excel/excel01.xlsx");
```

#### Sample 2
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

#### Sample 3
Getting a stream reader to a generated excel temp file that gets automatically deleted once the reader is closed.

```csharp
using Com.H.Excel;
var list = new List<object>() {
	new { Name = "John", Age = 20 },
	new { Name = "Jane", Age = 21 },
	new { Name = "Jack", Age = 22 }
};
var stream = list.ToExcelStream();
```


### Reading Excel
#### Sample 1
Reading a single sheet in an excel file

```csharp
using Com.H.Excel;

string filePath = @"c:/temp/excel/excel02.xlsx";

using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
	// Note: to parse to a pre-defined non dynamic class (e.g. Person) use:
    // var sheet = fileStream.ParseExcelSheet<Person>("Sheet1");
	
    var sheet = fileStream.ParseExcelSheet("Sheet1");
    foreach (var person in sheet)
    {
        Console.WriteLine($"name = {person.Name}, age = {person.Age}");
    }
}
```

#### Sample 2
Reading all sheets in an excel file

```csharp
using Com.H.Excel;

string filePath = @"c:/temp/excel/excel02.xlsx";

using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    var sheets = fileStream.ParseExcel();
    
    foreach(var sheet in sheets)
    {
        Console.WriteLine("-------------------");
        Console.WriteLine($"Sheet: {sheet.Key}");
        Console.WriteLine("-------------------");
        foreach (var person in sheet.Value)
        {
            Console.WriteLine($"name = {person.Name}, age = {person.Age}");
        }
        Console.WriteLine();
    }
}
```