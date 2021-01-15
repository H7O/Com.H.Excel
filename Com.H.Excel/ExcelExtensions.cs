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

        
        private static CellValues GetOpenXmlType(this Type type)
        {
            if (OpenXmlTypesDic.Keys.Contains(type))
                return
                    OpenXmlTypesDic[type];
            
            return CellValues.String;
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

    }
}
