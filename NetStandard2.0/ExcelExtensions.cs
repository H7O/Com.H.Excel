using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using System.Dynamic;

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

        /// <summary>
        /// https://github.com/closedxml/closedxml/wiki/NumberFormatId-Lookup-Table
        /// range.Style.NumberFormat.NumberFormatId = #;
        /// **ID**** Format Code**
        /// 0	General
        /// 1	0
        /// 2	0.00
        /// 3	#,##0
        /// 4	#,##0.00
        /// 9	0%
        /// 10	0.00%
        /// 11	0.00E+00
        /// 12	# ?/?
        /// 13	# ??/??
        /// 14	d/m/yyyy
        /// 15	d-mmm-yy
        /// 16	d-mmm
        /// 17	mmm-yy
        /// 18	h:mm tt
        /// 19	h:mm:ss tt
        /// 20	H:mm
        /// 21	H:mm:ss
        /// 22	m/d/yyyy H:mm
        /// 37	#,##0 ;(#,##0)
        /// 38	#,##0 ;[Red](#,##0)
        /// 39	#,##0.00;(#,##0.00)
        /// 40	#,##0.00;[Red](#,##0.00)
        /// 45	mm:ss
        /// 46	[h]:mm:ss
        /// 47	mmss.0
        /// 48	##0.0E+0
        /// 49	@
        /// </summary>
        private enum Formats
        {
            General = 0,
            Integer1 = 1,
            Decimal1 = 2,
            Integer2 = 3,
            Decimal2 = 4,
            IntegerPercentage = 9,
            DecimalPercentage = 10,
            DecimalScientific1 = 11,
            Date1 = 14,
            Date2 = 15,
            Date3 = 16,
            Time1 = 18,
            Time2 = 19,
            Time3 = 20,
            Time4 = 21,
            DateTime1 = 22,
            Integer3 = 37,
            Integer4 = 38,
            Decimal3 = 39,
            Decimal4 = 40,
            Time5 = 45,
            Time6 = 46,
            Time7 = 47,
            DecimalScientific2 = 48,
            Currency = 164,
            Accounting = 44,
            Date4 = 165,
            Time8 = 166,
            Fraction = 12,
            // Scientific = 11,
            Text = 49
        }

        private static readonly uint[] DateTimeFormatIds
            = new uint[]
            {
                (uint) Formats.Date1,
                (uint) Formats.Date2,
                (uint) Formats.Date3,
                (uint) Formats.DateTime1,
                (uint) Formats.Date4,
                (uint) Formats.Time1,
                (uint) Formats.Time2,
                (uint) Formats.Time3,
                (uint) Formats.Time4,
                (uint) Formats.Time5,
                (uint) Formats.Time6,
                (uint) Formats.Time7,
                (uint) Formats.Time8
            };
        private static readonly uint[] IntFormatIds
            = new uint[]
            {
                (uint) Formats.Integer1,
                (uint) Formats.Integer2,
                (uint) Formats.Integer3,
                (uint) Formats.Integer4,
                (uint) Formats.IntegerPercentage,
            };

        private static readonly uint[] DecimalFormatIds
            = new uint[]
            {
                (uint) Formats.Decimal1,
                (uint) Formats.Decimal2,
                (uint) Formats.Decimal3,
                (uint) Formats.Decimal4,
                (uint) Formats.DecimalPercentage,
                (uint) Formats.DecimalScientific1,
                (uint) Formats.DecimalScientific2
            };

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


        public static void WriteExcel<T>(
            this IDictionary<string, IEnumerable<T>> enumerables,
            Stream outStream,
            bool excludeHeaders = false
            ) where T : class
            =>
            WriteExcel(
                enumerables.ToDictionary(x => x.Key, v => v.Value.AsEnumerable<object>())
                , outStream
                , excludeHeaders
            );

        public static void WriteExcel<T>(
            this IDictionary<string, List<T>> enumerables,
            Stream outStream,
            bool excludeHeaders = false
            ) where T : class
            =>
            WriteExcel(
                enumerables.ToDictionary(x => x.Key, v => v.Value.AsEnumerable<object>())
                , outStream
                , excludeHeaders
            );

        public static void WriteExcel<T>(
            this IDictionary<string, IList<T>> enumerables,
            Stream outStream,
            bool excludeHeaders = false
            ) where T : class
            =>
            WriteExcel(
                enumerables.ToDictionary(x => x.Key, v => v.Value.AsEnumerable<object>())
                , outStream
                , excludeHeaders
            );

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
            Stylesheet styleSheet = new Stylesheet();

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

            CellStyles css = new CellStyles();
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
                    if (string.IsNullOrEmpty(enumerableSheet.Key)) 
                        throw new InvalidDataException("Empty sheet name in Excel is not allowed. Make sure the IDictionary<string, IEnumerable<object>> enumerables you're passing has non-empty and unique keys");
                    sheetNumber++;

                    // Add a WorksheetPart to the WorkbookPart.
                    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    Sheet sheet = new Sheet()
                    {
                        Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = (UInt32)sheetNumber,
                        Name = enumerableSheet.Key
                    };

                    sheets.Append(sheet);
                    if (enumerableSheet.Value is null) continue;
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
                    using (var inStream = File.OpenRead(tempPath))
                    {
                        inStream.CopyTo(outStream, 32000);
                        outStream.Flush();
                        inStream.Close();
                    }
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

        /// <summary>
        /// Reads an excel document and returns multiple excel sheets in the form of dictionary of dynamic collections, where the dictionary
        /// keys represent the sheet name and the dynamic list associated to the sheet name represent the data in a dynamic (ExpandoObject) form
        /// </summary>
        /// <param name="inStream">input stream to the excel document</param>
        /// <param name="noHeaders"></param>
        /// <returns>dictionary of dynamic collections, where the dictionary
        /// keys represent the sheet name and the dynamic list associated to the sheet name represent the data in a dynamic (ExpandoObject) form</returns>
        public static Dictionary<string, List<dynamic>> ParseExcel(
            this System.IO.Stream inStream, bool noHeaders = false)
        {
            SpreadsheetDocument doc = SpreadsheetDocument.Open(inStream, false);
            var excelSheets = doc?.WorkbookPart?.Workbook?
                    .Descendants<Sheet>();
            WorkbookPart workbookPart = doc.WorkbookPart;
            Dictionary<string, SheetData> sheets = excelSheets.ToDictionary(
                key => key.Name?.ToString(), value =>
                ((WorksheetPart)workbookPart.GetPartById(value.Id))
                    .Worksheet.GetFirstChild<SheetData>());

            Dictionary<string, List<dynamic>> result = new Dictionary<string, List<dynamic>>();

            foreach (var sheet in sheets)
            {
                result.Add(sheet.Key, new List<dynamic>());

                Dictionary<string, Type> headers = noHeaders ?
                         Enumerable.Range(0, sheet.Value.Count() - 1)
                        .Select(x => $"column_{x}")
                        .ToDictionary(key => key, value => typeof(string))
                        : sheet.Value?.FirstOrDefault()?.Select(x =>
                         ((Cell)x).GetText(workbookPart))
                        .ToDictionary(key => key, value => typeof(string));

                List<string> headerNames = headers.Keys?.ToList();


                foreach (Row row in sheet.Value.Skip(noHeaders ? 0 : 1).Cast<Row>())
                {
                    ExpandoObject d = new ExpandoObject();
                    int headerIndex = -1;
                    foreach (Cell cell in row.Select(x => (Cell)x))
                    {
                        headerIndex++;
                        var headerName = headerNames[headerIndex];
                        Type type =
                            (headers[headerName] = cell.GetDataTypeOtherThanString(workbookPart)
                            ?? headers[headerName] ?? typeof(string));

                        if (cell == null)
                        {
                            ((IDictionary<String, Object>)d)[headerName] = type.GetDefault();
                            continue;
                        }

                        int? index = (cell.GetCellColIndex() - 1)??headerIndex;
                        // if (index == null) break;

                        // fill collapsed columns
                        if (index > headerIndex)
                            headerName = headerNames[headerIndex = 
                                Enumerable.Range(headerIndex, (int)index - headerIndex)
                                .Aggregate(headerIndex, (i, n) =>
                                {
                                    ((IDictionary<String, Object>)d)[headerName] = 
                                        (headers[headerNames[i]] = cell.GetDataTypeOtherThanString(workbookPart)
                                        ?? headers[headerNames[i]] ?? typeof(string)).GetDefault();
                                    return n + 1;
                                })];

                        object value = null;

                        try
                        {
                            value = cell.GetObject(workbookPart); // Convert.ChangeType(cell.GetText(doc), type, CultureInfo.InvariantCulture);
                        }
                        catch { }
                        ((IDictionary<String, Object>)d)[headerName] = value ?? type.GetDefault();

                    }
                    result[sheet.Key].Add(d);
                }


            }
            #region finalazing
            doc.Close();

            #endregion
            return result;
        }

        public static List<T> ParseExcel<T>(
    this System.IO.Stream inStream, string sheetName = null, bool noHeaders = false)
        {
            SpreadsheetDocument doc = SpreadsheetDocument.Open(inStream, false);
            var excelSheets = doc?.WorkbookPart?.Workbook?
                    .Descendants<Sheet>();
            WorkbookPart workbookPart = doc.WorkbookPart;
            Dictionary<string, SheetData> sheets = excelSheets.ToDictionary(
                key => key.Name?.ToString(), value =>
                ((WorksheetPart)workbookPart.GetPartById(value.Id))
                    .Worksheet.GetFirstChild<SheetData>());
            var sheet = sheets.FirstOrDefault(x => x.Key.EqualsIgnoreCase(sheetName));

            if (sheet.Equals(default(KeyValuePair<string, SheetData>)))
                sheet = sheets.FirstOrDefault();

            if (sheet.Equals(default(KeyValuePair<string, SheetData>)))
                return null;
            int hIndex = 0;
            Dictionary<int, PropertyInfo> headers = noHeaders ?
                    typeof(T).GetCachedProperties()?
                    .ToDictionary(k => hIndex++, v => v.Info)
                     : sheet.Value.FirstOrDefault()?
                    .Select(x => ((Cell)x).GetText(workbookPart))
                    .LeftJoin(typeof(T).GetCachedProperties(),
                        e => e?.ToUpperInvariant(),
                        p => p.Name?.ToUpperInvariant(),
                        (e, p) => new { Index = hIndex++, p.Info}
                    ).ToDictionary(k => k.Index, v => v.Info);

            if (headers is null || headers.Count < 1) return null;


            List<T> result = new List<T>();
            var hCount = headers.Count;

            foreach (Row row in sheet.Value.Skip(noHeaders ? 0 : 1).Cast<Row>())
            {
                T d = Activator.CreateInstance<T>();
                result.Add(d);

                
                int index = -1;
                foreach (Cell cell in row.Select(x => (Cell)x))
                {
                    if ((index = (cell.GetCellColIndex() - 1) ?? ++index)
                        > hCount
                        && noHeaders) break;

                    var pInfo = headers[index];
                    if (pInfo is null) continue;

                    try
                    {
                        var value = cell.GetObject(workbookPart);
                        if (value == null)
                        {
                            pInfo.SetValue(d, pInfo.PropertyType.GetDefault());
                            continue;
                        }
                        pInfo.SetValue(d, value.ConvertTo(pInfo.PropertyType));
                    }
                    catch { }

                }
            }
            #region finalazing
            doc.Close();

            #endregion


            return result;
        }


        public static List<T> ParseExcelDepricated<T>(
            this System.IO.Stream inStream, string sheetName = null)
        {
            SpreadsheetDocument doc = SpreadsheetDocument.Open(inStream, false);
            var excelSheets = doc?.WorkbookPart?.Workbook?
                    .Descendants<Sheet>();
            WorkbookPart workbookPart = doc.WorkbookPart;
            var sheetId = doc?.WorkbookPart?.Workbook?
                    .Descendants<Sheet>().FirstOrDefault(x =>
                string.IsNullOrWhiteSpace(sheetName)
                || sheetName.EqualsIgnoreCase(x.Name))?.Id;

            if (sheetId == null) return null;
            var sheet = ((WorksheetPart)workbookPart.GetPartById(sheetId))?.Worksheet?.GetFirstChild<SheetData>();
            if (sheet == null) return null;

            var headers = sheet.FirstOrDefault()?
            .Select(x => ((Cell)x).GetText(workbookPart))
            .LeftJoin(typeof(T).GetCachedProperties(),
                e => e?.ToUpperInvariant(),
                t => t.Name?.ToUpperInvariant(),
                (e, p) => new { Excel = e, PInfo = p.Info }
            ).ToList();

            if (headers == null || headers.Count < 1) return null;

            List<T> result = new List<T>();
            foreach (Row row in sheet.Skip(1).Cast<Row>())
            {
                T obj = Activator.CreateInstance<T>();
                int lastIndex = 0;
                foreach (Cell cell in row.Select(x => (Cell)x))
                {
                    int? index = cell?.GetCellColIndex()??lastIndex - 1;

                    // fill collapsed columns
                    if (index > lastIndex)
                        lastIndex = Enumerable.Range(lastIndex, (int)index - lastIndex)
                                .Aggregate(lastIndex, (i, n) =>
                                {
                                    var pInfoInternal = headers[i]?.PInfo;
                                    if (pInfoInternal != null)
                                        pInfoInternal.SetValue(obj, pInfoInternal.PropertyType.GetDefault());
                                    return n + 1;
                                });


                    var pInfo = headers[lastIndex++].PInfo;
                    if (pInfo == null) continue;
                    try
                    {
                        var value = cell.GetObject(workbookPart);
                        if (value == null)
                        {
                            pInfo.SetValue(obj, pInfo.PropertyType.GetDefault());
                            continue;
                        }
                        pInfo.SetValue(obj, value.ConvertTo(pInfo.PropertyType));

                    }
                    catch { }
                    
                }
                result.Add(obj);
            }
            #region finalazing
            doc.Close();

            #endregion


            return result;
        }


        #region extract info
        private static Type GetDataTypeOtherThanString(this Cell cell, WorkbookPart workbookPart)
        {
            if (cell?.DataType?.Value != null) return cell.DataType.Value.GetTypeFromOpenXml();

            if (cell?.StyleIndex?.Value == null) return null;

            int? styleIndex = (int?)cell.StyleIndex?.Value;
            if (styleIndex == null) return null;
            CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt((int)styleIndex);
            uint formatId = cellFormat.NumberFormatId.Value;

            if (DateTimeFormatIds.Contains(formatId)) return typeof(DateTime?);
            if (IntFormatIds.Contains(formatId)) return typeof(int?);
            if (DecimalFormatIds.Contains(formatId)) return typeof(decimal?);
            return null;
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

            CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt((int)styleIndex);
            uint formatId = cellFormat.NumberFormatId.Value;
            if (DateTimeFormatIds.Contains(formatId) && double.TryParse(cell.InnerText, out double oaDate))
                return DateTime.FromOADate(oaDate);
            else if (DecimalFormatIds.Contains(formatId)
                && decimal.TryParse(cell.InnerText, out decimal d)
                ) return d;
            else if (IntFormatIds.Contains(formatId)
                && int.TryParse(cell.InnerText, out int i)) return i;

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

        #endregion



    }
}
