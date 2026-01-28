using Com.H.Excel;
using Xunit;

namespace Com.H.Excel.Tests;

public class ExcelTests : IDisposable
{
    private readonly string _tempFolder;

    public ExcelTests()
    {
        _tempFolder = Path.Combine(Path.GetTempPath(), "ComHExcelTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempFolder);
    }

    public void Dispose()
    {
        // Cleanup temp folder after tests
        if (Directory.Exists(_tempFolder))
        {
            try { Directory.Delete(_tempFolder, true); } catch { }
        }
    }

    #region Writing Tests

    [Fact]
    public void WriteSingleSheet_CreatesValidExcelFile()
    {
        // Arrange
        var list = new List<object>()
        {
            new { Name = "John", Age = 20 },
            new { Name = "Jane", Age = 21 },
            new { Name = "Jack", Age = 22 }
        };
        var filePath = Path.Combine(_tempFolder, "single_sheet.xlsx");

        // Act
        list.ToExcelFile(filePath);

        // Assert
        Assert.True(File.Exists(filePath));
        Assert.True(new FileInfo(filePath).Length > 0);
    }

    [Fact]
    public void WriteMultiSheet_CreatesValidExcelFile()
    {
        // Arrange
        var sheet1 = new List<object>()
        {
            new { Name = "John", Age = 20 },
            new { Name = "Jane", Age = 21 }
        };

        var sheet2 = new List<object>()
        {
            new { Name = "Tom", Age = 30 },
            new { Name = "Helen", Age = 31 }
        };

        var sheets = new Dictionary<string, IEnumerable<object>>()
        {
            { "Employees", sheet1 },
            { "Contractors", sheet2 }
        };

        var filePath = Path.Combine(_tempFolder, "multi_sheet.xlsx");

        // Act
        sheets.ToExcelFile(filePath);

        // Assert
        Assert.True(File.Exists(filePath));
        Assert.True(new FileInfo(filePath).Length > 0);
    }

    [Fact]
    public void WriteToStream_ReturnsValidStream()
    {
        // Arrange
        var list = new List<object>()
        {
            new { Product = "Apple", Price = 1.50 },
            new { Product = "Banana", Price = 0.75 }
        };

        // Act
        using var stream = list.ToExcelStream();

        // Assert
        Assert.NotNull(stream);
        Assert.True(stream.CanRead);
        Assert.True(stream.Length > 0);
    }

    #endregion

    #region Reading Tests

    [Fact]
    public void ReadSingleSheet_ReturnsCorrectData()
    {
        // Arrange - Create test file first
        var originalData = new List<object>()
        {
            new { Name = "Alice", Age = 25 },
            new { Name = "Bob", Age = 30 },
            new { Name = "Charlie", Age = 35 }
        };
        var filePath = Path.Combine(_tempFolder, "read_single.xlsx");
        originalData.ToExcelFile(filePath);

        // Act
        using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var sheet = fileStream.ParseExcelSheet("Sheet1").ToList();

        // Assert
        Assert.Equal(3, sheet.Count);
        Assert.Equal("Alice", (string)sheet[0].Name);
        Assert.Equal("25", (string)sheet[0].Age);
        Assert.Equal("Bob", (string)sheet[1].Name);
        Assert.Equal("Charlie", (string)sheet[2].Name);
    }

    [Fact]
    public void ReadAllSheets_ReturnsAllSheets()
    {
        // Arrange - Create multi-sheet test file
        var sheet1 = new List<object>()
        {
            new { City = "New York", Population = 8000000 },
            new { City = "Los Angeles", Population = 4000000 }
        };

        var sheet2 = new List<object>()
        {
            new { Country = "USA", Capital = "Washington" },
            new { Country = "Canada", Capital = "Ottawa" }
        };

        var sheets = new Dictionary<string, IEnumerable<object>>()
        {
            { "Cities", sheet1 },
            { "Countries", sheet2 }
        };

        var filePath = Path.Combine(_tempFolder, "read_multi.xlsx");
        sheets.ToExcelFile(filePath);

        // Act
        using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var result = fileStream.ParseExcel();

        // Assert
        Assert.Equal(2, result.Count);
        Assert.True(result.ContainsKey("Cities"));
        Assert.True(result.ContainsKey("Countries"));
        Assert.Equal(2, result["Cities"].Count());
        Assert.Equal(2, result["Countries"].Count());
    }

    [Fact]
    public void ReadTypedSheet_ReturnsTypedObjects()
    {
        // Arrange - Create test file
        var originalData = new List<object>()
        {
            new { Name = "Product1", Price = 10.5 },
            new { Name = "Product2", Price = 20.0 }
        };
        var filePath = Path.Combine(_tempFolder, "read_typed.xlsx");
        originalData.ToExcelFile(filePath);

        // Act
        using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var products = fileStream.ParseExcelSheet<Product>("Sheet1").ToList();

        // Assert
        Assert.Equal(2, products.Count);
        Assert.Equal("Product1", products[0].Name);
        Assert.Equal("Product2", products[1].Name);
    }

    #endregion

    #region Round-Trip Tests

    [Fact]
    public void WriteAndRead_DataIntegrity()
    {
        // Arrange
        var originalData = new List<object>()
        {
            new { Id = 1, Description = "First item", Value = 100.50 },
            new { Id = 2, Description = "Second item", Value = 200.75 },
            new { Id = 3, Description = "Third item", Value = 300.25 }
        };
        var filePath = Path.Combine(_tempFolder, "roundtrip.xlsx");

        // Act - Write
        originalData.ToExcelFile(filePath);

        // Act - Read
        using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var readData = fileStream.ParseExcelSheet("Sheet1").ToList();

        // Assert
        Assert.Equal(3, readData.Count);
        Assert.Equal("1", (string)readData[0].Id);
        Assert.Equal("First item", (string)readData[0].Description);
        Assert.Equal("Second item", (string)readData[1].Description);
        Assert.Equal("Third item", (string)readData[2].Description);
    }

    #endregion

    #region Empty/Null Cell Tests

    [Fact]
    public void ReadSheet_WithEmptyCellsInMiddle_HandlesCorrectly()
    {
        // This test verifies that empty cells in the middle of a row are handled correctly.
        // OpenXml doesn't serialize empty cells, so column indexing can skip over them.
        // The library should detect this gap and fill with null/default values.

        // Arrange - Create test file with empty cells
        // We'll write data where middle column has null to simulate empty cell
        var originalData = new List<object>()
        {
            new { ColA = "A1", ColB = "B1", ColC = "C1", ColD = "D1" },
            new { ColA = "A2", ColB = (string?)null, ColC = "C2", ColD = "D2" },  // B2 is null
            new { ColA = "A3", ColB = "B3", ColC = (string?)null, ColD = "D3" },  // C3 is null
            new { ColA = (string?)null, ColB = "B4", ColC = "C4", ColD = "D4" },  // A4 is null
        };
        var filePath = Path.Combine(_tempFolder, "empty_cells.xlsx");
        originalData.ToExcelFile(filePath);

        // Act
        using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var readData = fileStream.ParseExcelSheet("Sheet1").ToList();

        // Assert
        Assert.Equal(4, readData.Count);
        
        // Row 1 - all values present
        Assert.Equal("A1", (string)readData[0].ColA);
        Assert.Equal("B1", (string)readData[0].ColB);
        Assert.Equal("C1", (string)readData[0].ColC);
        Assert.Equal("D1", (string)readData[0].ColD);
        
        // Row 2 - B2 was null
        Assert.Equal("A2", (string)readData[1].ColA);
        Assert.Equal("C2", (string)readData[1].ColC);
        Assert.Equal("D2", (string)readData[1].ColD);
        
        // Row 3 - C3 was null
        Assert.Equal("A3", (string)readData[2].ColA);
        Assert.Equal("B3", (string)readData[2].ColB);
        Assert.Equal("D3", (string)readData[2].ColD);
        
        // Row 4 - A4 was null
        Assert.Equal("B4", (string)readData[3].ColB);
        Assert.Equal("C4", (string)readData[3].ColC);
        Assert.Equal("D4", (string)readData[3].ColD);
    }

    [Fact]
    public void ReadSheet_WithMultipleConsecutiveEmptyCells_HandlesCorrectly()
    {
        // Test multiple consecutive empty cells to ensure column index gap detection works
        
        // Arrange
        var originalData = new List<object>()
        {
            new { Col1 = "A", Col2 = "B", Col3 = "C", Col4 = "D", Col5 = "E" },
            new { Col1 = "X", Col2 = (string?)null, Col3 = (string?)null, Col4 = (string?)null, Col5 = "Y" },  // 3 consecutive nulls
        };
        var filePath = Path.Combine(_tempFolder, "consecutive_empty.xlsx");
        originalData.ToExcelFile(filePath);

        // Act
        using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var readData = fileStream.ParseExcelSheet("Sheet1").ToList();

        // Assert
        Assert.Equal(2, readData.Count);
        Assert.Equal("X", (string)readData[1].Col1);
        Assert.Equal("Y", (string)readData[1].Col5);
    }

    [Fact]
    public void ReadSheet_WithEmptyLastCells_HandlesCorrectly()
    {
        // Empty cells at the end of a row

        // Arrange
        var originalData = new List<object>()
        {
            new { First = "A", Second = "B", Third = "C" },
            new { First = "X", Second = (string?)null, Third = (string?)null },  // Last 2 cells empty
        };
        var filePath = Path.Combine(_tempFolder, "empty_last.xlsx");
        originalData.ToExcelFile(filePath);

        // Act
        using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var readData = fileStream.ParseExcelSheet("Sheet1").ToList();

        // Assert
        Assert.Equal(2, readData.Count);
        Assert.Equal("X", (string)readData[1].First);
    }

    #endregion
}

// Helper class for typed parsing test
public class Product
{
    public string? Name { get; set; }
    public double Price { get; set; }
}
