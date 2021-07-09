using Com.H.Reflection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Com.H.Xml;
using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using Com.H.IO;
using System.Data;
using System.Dynamic;
using Com.H.Text;

namespace Com.H.Excel
{
    public static class ExcelExtensions
    {
        #region OpenXml Excel type dictionary
        private static Dictionary<Type, CellValues> OpenXmlTypesDic { get; set; }
        = new Dictionary<Type, CellValues>()
        {
            {typeof(string),  CellValues.String},
            {typeof(int?), CellValues.Number},
            {typeof(decimal?), CellValues.Number },
            {typeof(double?), CellValues.Number },
            {typeof(float?), CellValues.Number },
            {typeof(long?), CellValues.Number },
            {typeof(bool?), CellValues.Boolean },
            {typeof(DateTime?), CellValues.Number },
            {typeof(int), CellValues.Number},
            {typeof(decimal), CellValues.Number },
            {typeof(double), CellValues.Number },
            {typeof(float), CellValues.Number },
            {typeof(long), CellValues.Number },
            {typeof(bool), CellValues.Boolean },
            {typeof(DateTime), CellValues.Number },
            };

        private static Dictionary<CellValues, Type> TypeToOpenXmlDic { get; set; }
        = new Dictionary<CellValues, Type>()
        {
            {CellValues.String, typeof(string) },
            {CellValues.Number, typeof(decimal?)},
            {CellValues.Boolean, typeof(bool?) },
            {CellValues.Date, typeof(DateTime?) }
            };


        private static CellValues GetOpenXmlType(this Type type)
        {
            if (OpenXmlTypesDic.Keys.Contains(type))
                return
                    OpenXmlTypesDic[type];

            return CellValues.String;
        }

        private static Type GetTypeFromOpenXml(this CellValues type)
        {
            if (TypeToOpenXmlDic.Keys.Contains(type))
                return
                    TypeToOpenXmlDic[type];

            return typeof(string);
        }


        #endregion


        #region format types
        private enum Formats
        {
            General = 0,
            Number = 1,
            Decimal = 2,
            Currency = 164,
            Accounting = 44,
            DateShort = 14,
            DateLong = 165,
            Time = 166,
            Percentage = 10,
            Fraction = 12,
            Scientific = 11,
            Text = 49
        }

        #endregion

        #region excel generation
        public static Stream ToExcelReader(
            this IDictionary<string, IEnumerable<object>> enumerables,
            string preferredTempFolderPath = null,
            string preferredTempFileName = null
            )
        {
            string excelOutputPath = enumerables.ToExcelTempFile(preferredTempFolderPath, preferredTempFileName);
            return new FileStream(excelOutputPath, FileMode.Open, FileAccess.Read, FileShare.None, 4096, FileOptions.DeleteOnClose);
        }

        public static string ToExcelTempFile(
            this IDictionary<string, IEnumerable<object>> enumerables,
            string preferredTempFolderPath = null,
            string preferredTempFileName = null
            )
        {
            string tempBasePath =
                Path.Combine(
                (string.IsNullOrEmpty(preferredTempFolderPath) ?
                Path.GetTempPath() : preferredTempFolderPath));

            var path =
                Path.Combine(tempBasePath,
                (string.IsNullOrEmpty(preferredTempFileName) ?
                Guid.NewGuid().ToString() + ".xlsx"
                : preferredTempFileName)).EnsureParentDirectory();


            if (File.Exists(path))
            {
                try
                {
                    File.Delete(path);
                }
                catch { }
            }
            using (StreamWriter f = File.CreateText(path))
            {
                enumerables.WriteExcel(f.BaseStream);
                f.Close();
            }
            return path;
        }


        public static void WriteExcel(
            this IDictionary<string, List<dynamic>> enumerables,
            Stream outStream,
            bool excludeHeaders = false
            )
            =>
            WriteExcel(
                enumerables.ToDictionary(x => x.Key, v => v.Value.AsEnumerable<object>())
                , outStream
                , excludeHeaders
            );

        public static void WriteExcel(
            this IDictionary<string, IList<dynamic>> enumerables,
            Stream outStream,
            bool excludeHeaders = false
            )
            =>
            WriteExcel(
                enumerables.ToDictionary(x => x.Key, v => v.Value.AsEnumerable<object>())
                , outStream
                , excludeHeaders
            );

        private static void AddDateStyle(this WorkbookPart workbookPart)
        {
            Stylesheet styleSheet = new();

            var cf1 = new CellFormat
            {
                NumberFormatId = 0,
                ApplyNumberFormat = true
            };

            var cf2 = new CellFormat
            {
                NumberFormatId = 15,
                ApplyNumberFormat = true
            };

            var cfs = new CellFormats();
            cfs.Append(cf1);
            cfs.Append(cf2);
            styleSheet.CellFormats = cfs;

            styleSheet.Borders = new Borders();
            styleSheet.Borders.Append(new Border());
            styleSheet.Fills = new Fills();
            styleSheet.Fills.Append(new Fill());
            styleSheet.Fonts = new Fonts();
            styleSheet.Fonts.Append(new Font());

            workbookPart.AddNewPart<WorkbookStylesPart>();
            workbookPart.WorkbookStylesPart.Stylesheet = styleSheet;

            CellStyles css = new();
            var cs = new CellStyle
            {
                FormatId = 0,
                BuiltinId = 0
            };
            css.Append(cs);
            css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
            styleSheet.Append(css);
        }

        public static void WriteExcel(
            this IDictionary<string, IEnumerable<object>> enumerables,
            System.IO.Stream outStream,
            bool excludeHeaders = false)
        {
            try
            {
                #region initial settings

                string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");

                SpreadsheetDocument spreadsheetDocument =
                    SpreadsheetDocument.Create(tempPath, SpreadsheetDocumentType.Workbook, true);

                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.AddDateStyle();
                workbookpart.Workbook = new Workbook();

                // Add Sheets to the Workbook.
                Sheets sheets =
                    spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());



                #endregion

                #region looping sheets

                var sheetNumber = 0;
                foreach (var enumerableSheet in enumerables)
                {
                    sheetNumber++;

                    // Add a WorksheetPart to the WorkbookPart.
                    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    Sheet sheet = new()
                    {
                        Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = (UInt32)sheetNumber,
                        Name = enumerableSheet.Key
                    };

                    sheets.Append(sheet);
                    if (enumerableSheet.Value == null) continue;

                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                    bool headersSet = false;

                    // looping data rows to fill excel sheet
                    foreach (var item in enumerableSheet.Value)
                    {
                        var properties = item?.GetCachedProperties()?.ToList();

                        if (properties == null || properties.Count < 1) continue;

                        #region headers
                        if (!excludeHeaders && !headersSet)
                        {
                            sheetData.Append(
                            new Row(
                            properties.Select(pInfo => new Cell()
                            {
                                CellValue = new CellValue(pInfo.Name),
                                DataType = new EnumValue<CellValues>(CellValues.String)
                            })));
                            headersSet = true;
                        }
                        #endregion


                        Cell GetCell((string Name, PropertyInfo Info) pInfo)
                        {
                            object valueRaw = pInfo.Info.GetValue(item);

                            if (valueRaw == null)
                                return new Cell()
                                {
                                    CellValue = new CellValue(""),
                                    DataType = new EnumValue<CellValues>(pInfo.Info.PropertyType.GetOpenXmlType())
                                };

                            if (pInfo.Info.PropertyType == typeof(DateTime)
                                ||
                                pInfo.Info.PropertyType == typeof(DateTime?))
                                return new Cell()
                                {
                                    CellValue = new CellValue(((DateTime)valueRaw)
                                    .ToOADate().ToString(CultureInfo.InvariantCulture)),
                                    DataType = new EnumValue<CellValues>(CellValues.Number),
                                    StyleIndex = 1
                                };

                            return new Cell()
                            {
                                CellValue = new CellValue(valueRaw.ToString()),
                                DataType = new EnumValue<CellValues>(
                                    pInfo.Info.PropertyType.GetOpenXmlType()),
                                StyleIndex = 0
                            };

                        }

                        #region data / filling cells (columns) within current row
                        sheetData.Append(new Row(properties
                            .Select(pInfo => GetCell(pInfo))));

                        #endregion


                    }

                }

                #endregion
                #region finalizing

                spreadsheetDocument.Close();

                spreadsheetDocument = null;
                try
                {
                    outStream.Write(File.ReadAllBytes(tempPath));
                    outStream.Flush();
                    File.Delete(tempPath);
                }
                catch { }
                #endregion
            }
            catch
            {
                throw;
            }
        }



        #endregion


        #region read excel

        public static Dictionary<string, List<dynamic>> ParseExcel(
            this System.IO.Stream inStream, bool skipHeaders = false)
        {
            SpreadsheetDocument doc = SpreadsheetDocument.Open(inStream, false);
            var excelSheets = doc?.WorkbookPart?.Workbook?
                    .Descendants<Sheet>();
            WorkbookPart workbookPart = doc.WorkbookPart;
            Dictionary<string, SheetData> sheets = excelSheets.ToDictionary(
                key => key.Name?.ToString(), value =>
                ((WorksheetPart)workbookPart.GetPartById(value.Id))
                    .Worksheet.GetFirstChild<SheetData>());

            Dictionary<string, List<dynamic>> result = new();

            foreach (var sheet in sheets)
            {
                result.Add(sheet.Key, new List<dynamic>());

                Dictionary<string, Type> headers = skipHeaders ?
                         Enumerable.Range(0, sheet.Value.Count() - 1)
                        .Select(x => $"column_{x}")
                        .ToDictionary(key => key, value => typeof(string))
                        : sheet.Value?.FirstOrDefault()?.Select(x =>
                         ((Cell)x).GetText(workbookPart)).ToDictionary(key => key, value => typeof(string));

                List<string> headerNames = headers.Keys?.ToList();


                foreach (Row row in sheet.Value.Skip(skipHeaders ? 0 : 1))
                {
                    ExpandoObject d = new();
                    int headerIndex = -1;
                    foreach (Cell cell in row.Select(x => (Cell)x))
                    {
                        headerIndex++;
                        var headerName = headerNames[headerIndex];
                        Type type =
                            (headers[headerName] = cell.GetDataType(workbookPart)
                            ?? headers[headerName]??typeof(string));

                        if (cell == null)
                        {
                            d.TryAdd(headerName, type.GetDefault());
                            continue;
                        }

                        int? index = cell.GetCellColIndex() - 1;
                        if (index == null) break;

                        // fill collapsed columns
                        if (index > headerIndex)
                            headerName = headerNames[headerIndex = Enumerable.Range(headerIndex, (int) index - headerIndex)
                                .Aggregate(headerIndex, (i, n) =>
                                {
                                    d.TryAdd(headerNames[i],
                                        (headers[headerNames[i]] = cell.GetDataType(workbookPart)
                                        ?? headers[headerNames[i]]??typeof(string)).GetDefault());
                                    return n+1;
                                })];

                        object value = null;

                        try
                        {
                            value = cell.GetObject(workbookPart); // Convert.ChangeType(cell.GetText(doc), type, CultureInfo.InvariantCulture);
                        }
                        catch { }
                        d.TryAdd(headerName, value ?? type.GetDefault());

                    }
                    result[sheet.Key].Add(d);
                }


            }
            return result;
        }
        private static Type GetDataType(this Cell cell, WorkbookPart workbookPart)
        {

            if (cell?.DataType?.Value != null) return cell.DataType.Value.GetTypeFromOpenXml();

            if (cell?.StyleIndex?.Value == null) return null;

            int? styleIndex = (int?)cell.StyleIndex?.Value;
            if (styleIndex == null) return null;
            CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt((int)styleIndex);
            uint formatId = cellFormat.NumberFormatId.Value;
            return formatId switch
            {
                (uint)Formats.DateShort => typeof(DateTime?),
                (uint)Formats.DateLong => typeof(DateTime?),
                (uint)Formats.Number => typeof(decimal?),
                (uint)Formats.Decimal => typeof(decimal?),
                _ => null
            };

        }
        private static string GetText(this Cell cell, WorkbookPart workbookPart)
        {
            string value =
            (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) ?
                workbookPart.SharedStringTablePart.SharedStringTable.ChildElements[int.Parse(cell.CellValue.Text)].InnerText :
                cell.CellValue.Text;
            return value;
        }

        private static object GetObject(this Cell cell, WorkbookPart workbookPart)
        {
            if (cell == null) return null;
            if (cell.DataType != null)
            {
                switch (cell.DataType?.Value)
                {
                    case CellValues.SharedString:
                        SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.CellValue.InnerText));
                        return ssi.Text.Text;
                    case CellValues.Boolean:
                        return cell.CellValue?.InnerText == "0";
                    default:
                        return cell.CellValue?.InnerText ?? cell.CellValue.Text;
                }
            }

            int? styleIndex = (int?)cell.StyleIndex?.Value;
            if (styleIndex == null) return cell.CellValue?.Text;
            
            CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt((int) styleIndex);
            uint formatId = cellFormat.NumberFormatId.Value;
            if (formatId == (uint)Formats.DateShort || formatId == (uint)Formats.DateLong)
            {
                if (double.TryParse(cell.InnerText, out double oaDate))
                    return DateTime.FromOADate(oaDate);
            }
            else
            {
                if (decimal.TryParse(cell.InnerText, out decimal d)) return d;
            }
            return cell.CellValue?.Text;
        }



        /// <summary>
        /// Extract column index from cell reference.
        /// edited from http://stackoverflow.com/questions/848147/how-to-convert-excel-sheet-column-names-into-numbers
        /// </summary>
        /// <param name="cell"></param>
        /// <returns>Cell column index</returns>
        /// 
        private static int? GetCellColIndex(this Cell cell)
            => cell?.CellReference?.Value?.ExtractAlphabet()?.ToUpper()?.
               Aggregate(0, (i, n) => 26 * i + n - 'A' + 1);


        #endregion
    }
}
