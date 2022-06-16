using CsvHelper;
using CsvHelper.Configuration;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLHelper
{
    internal class OpenXMLHelperWriter
    {
        private const int MaxRowValue = 1048576;
        private SharedStringTable SharedStringTable { get; set; }

        #region CreateCSV

        public static void CreateCSVFile<T>(string fileName, T[] record, bool autoMap)
        {
            using var ms = new MemoryStream();
            using StreamWriter sw = new(ms);
            CreateCSVFileObject(sw, ref record, autoMap);

            var result = Encoding.UTF8.GetString(ms.ToArray());
            File.WriteAllText(fileName, result);
        }

        public static void CreateCSVFile<T>(string fileName, List<T> record, bool autoMap)
        {
            using var ms = new MemoryStream();
            using StreamWriter sw = new(ms);
            CreateCSVFileObject(sw, ref record, autoMap);

            var result = Encoding.UTF8.GetString(ms.ToArray());
            File.WriteAllText(fileName, result);
        }

        public static byte[] CreateCSVFileToByteArray<T>(T[] record, bool autoMap)
        {
            using var ms = new MemoryStream();
            using StreamWriter sw = new(ms);
            CreateCSVFileObject(sw, ref record, autoMap);

            return ms.ToArray();
        }

        public static byte[] CreateCSVFileToByteArray<T>(List<T> record, bool autoMap)
        {
            using var ms = new MemoryStream();
            using StreamWriter sw = new(ms);
            CreateCSVFileObject(sw, ref record, autoMap);

            return ms.ToArray();
        }

        #endregion

        public static void CreateDocumentList<T>(SpreadsheetDocument spreadsheetDocument, List<T> records, bool header)
        {
            List<OpenXmlAttribute> oxa;
            OpenXmlWriter oxw;

            spreadsheetDocument.AddWorkbookPart();
            WorksheetPart wsp = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

            oxw = OpenXmlWriter.Create(wsp);
            oxw.WriteStartElement(new Worksheet());
            oxw.WriteStartElement(new SheetData());

            var propertyCount = records.First().GetType().GetProperties().Length;
            var rowCount = records.Count;
            var actualRow = 1;
            for (int i = 0; i < rowCount;)
            {
                oxa = new List<OpenXmlAttribute>();
                oxa.Add(new OpenXmlAttribute("r", null, (actualRow).ToString()));
                oxw.WriteStartElement(new Row(), oxa);
                var record = records[i];

                if (header)
                {
                    for (var j = 0; j < propertyCount; j++)
                    {
                        oxa = new()
                        {
                            // this is the data type ("t"), with CellValues.String ("str")
                            new OpenXmlAttribute("t", null, "str")
                        };
                        var value = GetPropertyName(record, j);
                        oxw.WriteStartElement(new Cell(), oxa);
                        //var value = feild.GetValue(recordsss, null);
                        oxw.WriteElement(new CellValue(value));
                        // this is for Cell
                        oxw.WriteEndElement();
                    }
                    header = false;
                }
                else
                {
                    for (var j = 0; j < propertyCount; j++)
                    {
                        oxa = new()
                        {
                            // this is the data type ("t"), with CellValues.String ("str")
                            new OpenXmlAttribute("t", null, "str")
                        };
                        var value = GetPropertyValue(record, j);
                        oxw.WriteStartElement(new Cell(), oxa);
                        //var value = feild.GetValue(recordsss, null);
                        oxw.WriteElement(new CellValue(value));
                        // this is for Cell
                        oxw.WriteEndElement();
                    }
                    i++;
                }
                oxw.WriteEndElement();
                actualRow++;
            }

            // this is for SheetData
            oxw.WriteEndElement();
            // this is for Worksheet
            oxw.WriteEndElement();
            oxw.Close();

            oxw = OpenXmlWriter.Create(spreadsheetDocument.WorkbookPart);
            oxw.WriteStartElement(new Workbook());
            oxw.WriteStartElement(new Sheets());

            // you can use object initialisers like this only when the properties
            // are actual properties. SDK classes sometimes have property-like properties
            // but are actually classes. For example, the Cell class has the CellValue
            // "property" but is actually a child class internally.
            // If the properties correspond to actual XML attributes, then you're fine.
            oxw.WriteElement(new Sheet()
            {
                Name = "Sheet1",
                SheetId = 1,
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(wsp)
            });

            // this is for Sheets
            oxw.WriteEndElement();
            // this is for Workbook
            oxw.WriteEndElement();
            oxw.Close();
        }

        public static void CreateDocumentArray<T>(SpreadsheetDocument spreadsheetDocument, T[] records, bool header)
        {
            List<OpenXmlAttribute> oxa;
            OpenXmlWriter oxw;

            spreadsheetDocument.AddWorkbookPart();
            WorksheetPart wsp = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

            oxw = OpenXmlWriter.Create(wsp);
            oxw.WriteStartElement(new Worksheet());
            oxw.WriteStartElement(new SheetData());

            var propertyCount = records.First().GetType().GetProperties().Length;
            var rowCount = records.Length;
            var actualRow = 1;
            for (int i = 0; i < rowCount;)
            {
                oxa = new List<OpenXmlAttribute>();
                oxa.Add(new OpenXmlAttribute("r", null, (actualRow).ToString()));
                oxw.WriteStartElement(new Row(), oxa);
                var record = records[i];

                if (header)
                {
                    for (var j = 0; j < propertyCount; j++)
                    {
                        oxa = new()
                        {
                            // this is the data type ("t"), with CellValues.String ("str")
                            new OpenXmlAttribute("t", null, "str")
                        };
                        var value = GetPropertyName(record, j);
                        oxw.WriteStartElement(new Cell(), oxa);
                        //var value = feild.GetValue(recordsss, null);
                        oxw.WriteElement(new CellValue(value));
                        // this is for Cell
                        oxw.WriteEndElement();
                    }
                    header = false;
                }
                else
                {
                    for (var j = 0; j < propertyCount; j++)
                    {
                        oxa = new()
                        {
                            // this is the data type ("t"), with CellValues.String ("str")
                            new OpenXmlAttribute("t", null, "str")
                        };
                        var value = GetPropertyValue(record, j);
                        oxw.WriteStartElement(new Cell(), oxa);
                        //var value = feild.GetValue(recordsss, null);
                        oxw.WriteElement(new CellValue(value));
                        // this is for Cell
                        oxw.WriteEndElement();
                    }
                    i++;
                }
                oxw.WriteEndElement();
                actualRow++;
            }

            // this is for SheetData
            oxw.WriteEndElement();
            // this is for Worksheet
            oxw.WriteEndElement();
            oxw.Close();

            oxw = OpenXmlWriter.Create(spreadsheetDocument.WorkbookPart);
            oxw.WriteStartElement(new Workbook());
            oxw.WriteStartElement(new Sheets());

            // you can use object initialisers like this only when the properties
            // are actual properties. SDK classes sometimes have property-like properties
            // but are actually classes. For example, the Cell class has the CellValue
            // "property" but is actually a child class internally.
            // If the properties correspond to actual XML attributes, then you're fine.
            oxw.WriteElement(new Sheet()
            {
                Name = "Sheet1",
                SheetId = 1,
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(wsp)
            });

            // this is for Sheets
            oxw.WriteEndElement();
            // this is for Workbook
            oxw.WriteEndElement();
            oxw.Close();
        }

        public static void CreateWorkbookUsingOpenXMLWriter<T>(List<T> records, string filePath, bool header)
        {
            var ms = new MemoryStream();
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
            CreateDocumentList(spreadsheetDocument, records, header);
            spreadsheetDocument.SaveAs(filePath);
            spreadsheetDocument.Close();
        }

        public static void CreateWorkbookUsingOpenXMLWriter<T>(T[] records, string filePath, bool header)
        {
            var ms = new MemoryStream();
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
            CreateDocumentArray(spreadsheetDocument, records, header);
            spreadsheetDocument.SaveAs(filePath);
            spreadsheetDocument.Close();
        }

        public static byte[] CreateWorkbookUsingOpenXMLWriterToByteArray<T>(List<T> records, bool header)
        {
            var ms = new MemoryStream();
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
            CreateDocumentList(spreadsheetDocument, records, header);
            spreadsheetDocument.Close();
            return ms.ToArray();
        }

        public static byte[] CreateWorkbookUsingOpenXMLWriterToByteArray<T>(T[] records, bool header)
        {
            var ms = new MemoryStream();
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
            CreateDocumentArray(spreadsheetDocument, records, header);
            spreadsheetDocument.Close();
            return ms.ToArray();
        }

        private static string GetPropertyName(object record, int count)
        {
            var value = "";
            var properties = record.GetType().GetProperties();

            var property = properties.ElementAt(count);

            var type = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
            if (type == typeof(string))
            {
                value = property.Name;
            }
            else
            {
                value = property.Name;
            }

            return value;
        }


        private static string GetPropertyValue(object record, int count)
        {
            var value = "";
            var properties = record.GetType().GetProperties();

            var property = properties.ElementAt(count);

            var type = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
            if (type == typeof(string))
            {
                value = property.GetValue(record, null).ToString();
            }
            else
            {
                value = property.GetValue(record, null).ToString();
            }

            return value;
        }

        #region CSV
        private static void CreateCSVFileObject<T>(StreamWriter sw, ref List<T> record, bool autoMap)
        {
            var config = new CsvConfiguration(CultureInfo.CurrentCulture) { Delimiter = ",", HasHeaderRecord = autoMap };
            using var csvWriter = new CsvWriter(sw, config);

            if (record.GetType() == typeof(List<string>))
            {
                foreach (var item in record)
                {
                    csvWriter.WriteField(item);
                    csvWriter.NextRecord();
                }
            }
            else
            {
                if (autoMap)
                {
                    csvWriter.WriteHeader<T>();
                    csvWriter.NextRecord();
                }
                csvWriter.WriteRecords(record);
            }
            sw.Flush();
        }

        private static void CreateCSVFileObject<T>(StreamWriter sw, ref T[] record, bool autoMap)
        {
            var config = new CsvConfiguration(CultureInfo.CurrentCulture) { Delimiter = ",", HasHeaderRecord = autoMap };
            using var csvWriter = new CsvWriter(sw, config);

            if (record.GetType() == typeof(List<string>))
            {
                foreach (var item in record)
                {
                    csvWriter.WriteField(item);
                    csvWriter.NextRecord();
                }
            }
            else
            {
                if (autoMap)
                {
                    csvWriter.WriteHeader<T>();
                    csvWriter.NextRecord();
                }
                csvWriter.WriteRecords(record);
            }
            sw.Flush();
        }
        #endregion
    }

}
