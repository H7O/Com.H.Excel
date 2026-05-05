using Com.H.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
        // Numeric cells now come back as int/decimal (was string in 10.1.x and earlier).
        Assert.Equal(25, (int)sheet[0].Age);
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
        Assert.Equal(1, (int)readData[0].Id);
        Assert.Equal(100.50m, (decimal)readData[0].Value);
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

    #region SharedString / InlineString edge cases

    // Builds a minimal XLSX with one sheet, one row, where the cell carries
    // DataType=SharedString but the workbook intentionally has NO SharedStringTablePart.
    // This mirrors malformed/unusual files some external writers produce and was
    // the original NRE the library hit.
    private static void BuildXlsxWithMissingSst(string path, string sharedStringIndex)
    {
        using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        // Row 1 header (inline string), Row 2 data cell pointing at SharedString index that doesn't resolve.
        var sheetData = new SheetData(
            new Row(new Cell
            {
                CellValue = new CellValue("Header"),
                DataType = new EnumValue<CellValues>(CellValues.String)
            }),
            new Row(new Cell
            {
                CellValue = new CellValue(sharedStringIndex),
                DataType = new EnumValue<CellValues>(CellValues.SharedString)
            })
        );
        worksheetPart.Worksheet = new Worksheet(sheetData);

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        });
        workbookPart.Workbook.Save();
    }

    private static void BuildXlsxWithInlineString(string path, string headerName, string inlineValue)
    {
        using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        var headerCell = new Cell
        {
            CellValue = new CellValue(headerName),
            DataType = new EnumValue<CellValues>(CellValues.String)
        };
        var dataCell = new Cell
        {
            DataType = new EnumValue<CellValues>(CellValues.InlineString),
            InlineString = new InlineString(new Text(inlineValue))
        };

        worksheetPart.Worksheet = new Worksheet(new SheetData(
            new Row(headerCell),
            new Row(dataCell)
        ));

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        });
        workbookPart.Workbook.Save();
    }

    [Fact]
    public void ReadSheet_WithMissingSharedStringTable_DoesNotThrow()
    {
        var filePath = Path.Combine(_tempFolder, "missing_sst.xlsx");
        BuildXlsxWithMissingSst(filePath, "0");

        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var rows = fs.ParseExcelSheet("Sheet1").ToList();

        Assert.Single(rows);
        // With no SST to resolve against, the raw index value is returned rather than NREing.
        Assert.Equal("0", (string)rows[0].Header);
    }

    [Fact]
    public void ReadSheet_WithMissingSharedStringTable_NonNumericIndex_DoesNotThrow()
    {
        var filePath = Path.Combine(_tempFolder, "missing_sst_garbled.xlsx");
        BuildXlsxWithMissingSst(filePath, "garbage-not-an-int");

        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var rows = fs.ParseExcelSheet("Sheet1").ToList();

        Assert.Single(rows);
        Assert.Equal("garbage-not-an-int", (string)rows[0].Header);
    }

    [Fact]
    public void ReadSheet_WithInlineString_ReturnsValue()
    {
        var filePath = Path.Combine(_tempFolder, "inline_string.xlsx");
        BuildXlsxWithInlineString(filePath, "Name", "Inline-Hello");

        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var rows = fs.ParseExcelSheet("Sheet1").ToList();

        Assert.Single(rows);
        Assert.Equal("Inline-Hello", (string)rows[0].Name);
    }

    #endregion

    #region Boolean round-trip

    public class BoolRow
    {
        public string? Name { get; set; }
        public bool Active { get; set; }
    }

    [Fact]
    public void WriteAndRead_Booleans_RoundTripCorrectly()
    {
        // Boolean cells in OpenXml use "1" for true and "0" for false.
        // Earlier the writer emitted "True"/"False" and the reader inverted the comparison;
        // both bugs combined could give wrong-but-consistent results, but any external tool
        // touching the file would see incorrect values. This test pins down the spec-compliant behavior.
        var data = new List<object>
        {
            new BoolRow { Name = "row-true",  Active = true  },
            new BoolRow { Name = "row-false", Active = false },
        };

        var filePath = Path.Combine(_tempFolder, "bool_roundtrip.xlsx");
        data.ToExcelFile(filePath);

        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var read = fs.ParseExcelSheet<BoolRow>("Sheet1").ToList();

        Assert.Equal(2, read.Count);
        Assert.Equal("row-true", read[0].Name);
        Assert.True(read[0].Active);
        Assert.Equal("row-false", read[1].Name);
        Assert.False(read[1].Active);
    }

    #endregion

    #region Null numeric / bool cells produce valid output

    public class NumericNullableRow
    {
        public string? Name { get; set; }
        public int? Score { get; set; }
        public DateTime? When { get; set; }
        public bool? Flag { get; set; }
    }

    [Fact]
    public void WriteAndRead_NullNumericAndDateAndBool_DoesNotCorruptFile()
    {
        // Previously the writer emitted <c t="n"><v></v></c> for null numerics — invalid per spec
        // and rejected by strict validators. This test asserts the file is well-formed enough
        // to round-trip and that nulls survive as defaults.
        var data = new List<object>
        {
            new NumericNullableRow { Name = "complete", Score = 42, When = new DateTime(2020,1,15), Flag = true },
            new NumericNullableRow { Name = "all-null", Score = null, When = null, Flag = null },
        };

        var filePath = Path.Combine(_tempFolder, "null_numeric.xlsx");
        data.ToExcelFile(filePath);

        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var read = fs.ParseExcel();

        Assert.True(read.ContainsKey("Sheet1"));
        var rows = read["Sheet1"];
        Assert.Equal(2, rows.Count);

        // First row has values
        Assert.Equal("complete", (string)rows[0].Name);

        // Second row's nulls should not blow up the parser.
        Assert.Equal("all-null", (string)rows[1].Name);
    }

    [Fact]
    public void WriteAndRead_NullNumeric_StillProducesValidOpenXml()
    {
        // Open the produced file via OpenXml's own Validator-like API to be sure it isn't malformed.
        var data = new List<object>
        {
            new NumericNullableRow { Name = "x", Score = null, When = null, Flag = null },
        };
        var filePath = Path.Combine(_tempFolder, "null_validation.xlsx");
        data.ToExcelFile(filePath);

        // If the file is malformed in a way that breaks OpenXml SDK reading,
        // SpreadsheetDocument.Open itself will throw. This is a structural sanity check.
        using var doc = SpreadsheetDocument.Open(filePath, false);
        Assert.NotNull(doc.WorkbookPart);
        Assert.NotNull(doc.WorkbookPart!.Workbook);
    }

    #endregion

    #region openpyxl-generated fixture regressions

    // openpyxl, by default, writes string cells as t="inlineStr" and emits no
    // sharedStrings.xml part at all. Reading those files used to NRE here on
    // workbookPart.SharedStringTablePart. Fixtures committed under TestFixtures/
    // pin down that exact byte pattern.

    private static string FixturePath(string fileName)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "TestFixtures", fileName);
        if (!File.Exists(path))
            throw new FileNotFoundException(
                $"Fixture not found at {path}. Regenerate by running TestFixtures/generate_openpyxl_fixtures.py via uv.");
        return path;
    }

    [Fact]
    public void ReadOpenpyxlFile_BasicStrings_ReturnsAllRows()
    {
        using var fs = File.OpenRead(FixturePath("openpyxl_basic_strings.xlsx"));
        var rows = fs.ParseExcelSheet().ToList();

        Assert.Equal(3, rows.Count);
        Assert.Equal("Alice", (string)rows[0].Name);
        Assert.Equal("Paris", (string)rows[0].City);
        Assert.Equal("Bob", (string)rows[1].Name);
        Assert.Equal("Tokyo", (string)rows[2].City);
    }

    [Fact]
    public void ReadOpenpyxlFile_MixedTypes_RoundTripsValues()
    {
        using var fs = File.OpenRead(FixturePath("openpyxl_mixed_types.xlsx"));
        var rows = fs.ParseExcel();

        Assert.True(rows.ContainsKey("Sheet1"));
        var data = rows["Sheet1"];
        Assert.Equal(2, data.Count);
        Assert.Equal("alice", (string)data[0].Name);
        Assert.Equal("bob", (string)data[1].Name);
    }

    [Fact]
    public void ReadOpenpyxlFile_EmptyMiddleCells_HandlesGapsCorrectly()
    {
        using var fs = File.OpenRead(FixturePath("openpyxl_empty_middle.xlsx"));
        var rows = fs.ParseExcelSheet().ToList();

        Assert.Equal(2, rows.Count);
        Assert.Equal("a1", (string)rows[0].A);
        Assert.Equal("c1", (string)rows[0].C);
        Assert.Equal("d1", (string)rows[0].D);
        Assert.Equal("a2", (string)rows[1].A);
        Assert.Equal("b2", (string)rows[1].B);
        Assert.Equal("d2", (string)rows[1].D);
    }

    [Fact]
    public void ReadOpenpyxlFile_DatesWithCustomNumFmt_ReturnsDateTimeValues()
    {
        // openpyxl writes date cells as <c t="n" s="N"><v>OADate</v></c> where the style
        // points at a CUSTOM numFmtId (typically 164) defined in the file's <numFmts>
        // table. Built-in date IDs aren't used. Reading these used to return the raw
        // OADate as a string because the library only consulted hardcoded format IDs.
        using var fs = File.OpenRead(FixturePath("openpyxl_dates_custom_numfmt.xlsx"));
        var rows = fs.ParseExcelSheet().ToList();

        Assert.Equal(2, rows.Count);

        Assert.IsType<DateTime>(rows[0].When);
        Assert.IsType<DateTime>(rows[1].When);

        Assert.Equal(new DateTime(2024, 3, 15, 9, 30, 0), (DateTime)rows[0].When);
        Assert.Equal(new DateTime(2025, 12, 31, 23, 59, 59),
            (DateTime)rows[1].When,
            TimeSpan.FromSeconds(1)); // OADate round-trip can drift sub-second
    }

    public class TimedRow
    {
        public string? Name { get; set; }
        public DateTime When { get; set; }
    }

    [Fact]
    public void ReadOpenpyxlFile_DatesWithCustomNumFmt_TypedParse()
    {
        using var fs = File.OpenRead(FixturePath("openpyxl_dates_custom_numfmt.xlsx"));
        var rows = fs.ParseExcelSheet<TimedRow>("Sheet1").ToList();

        Assert.Equal(2, rows.Count);
        Assert.Equal("alpha", rows[0].Name);
        Assert.Equal(new DateTime(2024, 3, 15, 9, 30, 0), rows[0].When);
    }

    [Fact]
    public void ReadOpenpyxlFile_PlainNumbers_ReturnAsNumericTypes()
    {
        // openpyxl writes numeric cells as <c t="n"><v>...</v></c> with NO style at all.
        // Previously these came back as strings; now they come back as int/decimal so
        // dynamic consumers see the value's actual type.
        using var fs = File.OpenRead(FixturePath("openpyxl_mixed_types.xlsx"));
        var rows = fs.ParseExcelSheet().ToList();

        Assert.Equal(2, rows.Count);
        Assert.IsType<int>(rows[0].Score);
        Assert.Equal(42, (int)rows[0].Score);
        Assert.Equal(7, (int)rows[1].Score);
    }

    #endregion

    #region Numeric handling — plain cells, integer-format with fractional values, locale

    [Fact]
    public void DynamicParse_PlainIntegerWrittenByUs_ComesBackAsInt()
    {
        var data = new List<object>
        {
            new { Name = "alpha", Count = 42 },
            new { Name = "beta", Count = 7 },
        };
        var path = Path.Combine(_tempFolder, "plain_int.xlsx");
        data.ToExcelFile(path);

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var rows = fs.ParseExcelSheet().ToList();

        Assert.Equal(2, rows.Count);
        Assert.IsType<int>(rows[0].Count);
        Assert.Equal(42, (int)rows[0].Count);
    }

    [Fact]
    public void DynamicParse_PlainDecimalWrittenByUs_ComesBackAsDecimal()
    {
        var data = new List<object>
        {
            new { Name = "alpha", Price = 10.5m },
        };
        var path = Path.Combine(_tempFolder, "plain_decimal.xlsx");
        data.ToExcelFile(path);

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var rows = fs.ParseExcelSheet().ToList();

        Assert.IsType<decimal>(rows[0].Price);
        Assert.Equal(10.5m, (decimal)rows[0].Price);
    }

    public class PricedRow
    {
        public string? Name { get; set; }
        public decimal Price { get; set; }
    }

    [Fact]
    public void IntegerFormatWithFractionalValue_FallsBackToDecimal()
    {
        // Build a file where the cell carries an integer-style format (#,##0)
        // but the underlying numeric value is fractional. Old code returned a
        // bare string because int.TryParse failed on "1234.5"; now we fall
        // through to decimal so no precision is lost.
        var path = Path.Combine(_tempFolder, "int_format_frac_value.xlsx");
        using (var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();

            var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet(
                new NumberingFormats(
                    new NumberingFormat { NumberFormatId = 164u, FormatCode = "#,##0" }),
                new Fonts(new Font()) { Count = 1 },
                new Fills(new Fill()) { Count = 1 },
                new Borders(new Border()) { Count = 1 },
                new CellFormats(
                    new CellFormat(),
                    new CellFormat
                    {
                        NumberFormatId = 164u,
                        ApplyNumberFormat = true
                    }) { Count = 2 });
            stylesPart.Stylesheet.Save();

            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData(
                new Row(
                    new Cell { CellValue = new CellValue("Name"), DataType = new EnumValue<CellValues>(CellValues.String) },
                    new Cell { CellValue = new CellValue("Score"), DataType = new EnumValue<CellValues>(CellValues.String) }),
                new Row(
                    new Cell { CellValue = new CellValue("alpha"), DataType = new EnumValue<CellValues>(CellValues.String) },
                    new Cell { CellValue = new CellValue("1234.5"), DataType = new EnumValue<CellValues>(CellValues.Number), StyleIndex = 1 })));

            var sheets = wbPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1, Name = "Sheet1" });
            wbPart.Workbook.Save();
        }

        using var fs = File.OpenRead(path);
        var rows = fs.ParseExcelSheet().ToList();

        Assert.Single(rows);
        // Format said int, but the value is fractional — must come back as decimal,
        // not as a bare string (would lose typing) and not as a truncated int (would lose precision).
        Assert.IsType<decimal>(rows[0].Score);
        Assert.Equal(1234.5m, (decimal)rows[0].Score);
    }

    [Fact]
    public void TypedParse_DecimalProperty_WorksRegardlessOfThreadCulture()
    {
        // Until 10.2.0, ConvertTo used current-culture for Convert.ChangeType. On a
        // de-DE machine, Convert.ChangeType("10.5", typeof(decimal)) parses the dot
        // as a thousands separator and returns 105m — silently corrupting values.
        // We force a de-DE culture for the duration of the test to lock down the fix.
        var data = new List<object>
        {
            new { Name = "alpha", Price = 10.5m },
            new { Name = "beta", Price = 1234.56m },
        };
        var path = Path.Combine(_tempFolder, "decimal_culture.xlsx");
        data.ToExcelFile(path);

        var prevCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
        try
        {
            System.Threading.Thread.CurrentThread.CurrentCulture =
                new System.Globalization.CultureInfo("de-DE");

            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var rows = fs.ParseExcelSheet<PricedRow>("Sheet1").ToList();

            Assert.Equal(2, rows.Count);
            Assert.Equal(10.5m, rows[0].Price);
            Assert.Equal(1234.56m, rows[1].Price);
        }
        finally
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = prevCulture;
        }
    }

    [Fact]
    public void DynamicParse_NumberInExtremelyLargeRange_FallsBackToDecimal()
    {
        // int.TryParse fails on values outside Int32 range — must fall back to decimal
        // rather than dropping to string.
        var path = Path.Combine(_tempFolder, "big_number.xlsx");
        using (var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData(
                new Row(new Cell { CellValue = new CellValue("Big"), DataType = new EnumValue<CellValues>(CellValues.String) }),
                new Row(new Cell { CellValue = new CellValue("9999999999999"), DataType = new EnumValue<CellValues>(CellValues.Number) })));
            var sheets = wbPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1, Name = "Sheet1" });
            wbPart.Workbook.Save();
        }

        using var fs = File.OpenRead(path);
        var rows = fs.ParseExcelSheet().ToList();
        Assert.Single(rows);
        Assert.IsType<decimal>(rows[0].Big);
        Assert.Equal(9999999999999m, (decimal)rows[0].Big);
    }

    #endregion
}

// Helper class for typed parsing test
public class Product
{
    public string? Name { get; set; }
    public double Price { get; set; }
}
