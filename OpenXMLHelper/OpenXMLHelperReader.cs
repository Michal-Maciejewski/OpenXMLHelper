using CsvHelper;
using CsvHelper.Configuration;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Dynamic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OpenXMLHelper
{
    internal class OpenXMLHelperReader
    {

        internal static List<object> GetListCSV(StreamReader streamReader, bool header)
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = ",", HasHeaderRecord = header };
            using var csvReader = new CsvReader(streamReader, config);
            return csvReader.GetRecords<object>().ToList();
        }

        internal static List<object> GetListExcelWorkbook(Stream stream, bool header)
        {
            using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
            var listObject = ReadExcelWorkbookData(spreadsheetDocument, header);
            return listObject;
        }

        private static List<object> ReadExcelWorkbookData(SpreadsheetDocument spreadsheetDocument, bool header)
        {
            var listObjects = new List<object>();
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SharedStringTable sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;

            OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
            var columnCount = 0;
            string[] columnName = Array.Empty<string>();
            int[] columnNumber = Array.Empty<int>();

            if (header)
            {
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(Row))
                    {
                        var row = (Row)reader.LoadCurrentElement();
                        var cells = row.Elements<Cell>().ToArray();
                        columnCount = cells.Length;
                        columnName = new string[columnCount];
                        columnNumber = new int[columnCount];

                        for (var count = 0; count < columnCount; count++)
                        {
                            var cell = cells[count];
                            var value = GetCellValue(ref cell, ref sharedStringTable);
                            columnNumber[count] = GetColumnIndex(cell.CellReference);
                            columnName[count] = value;
                        }
                        break;
                    }
                }
            }
            else
            {
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(Row))
                    {
                        var row = (Row)reader.LoadCurrentElement();
                        var cells = row.Elements<Cell>().ToArray();
                        columnCount = cells.Length;
                        columnName = new string[columnCount];
                        columnNumber = new int[columnCount];
                        var properties = new ExpandoObject() as IDictionary<string, Object>;

                        for (var count = 0; count < columnCount; count++)
                        {
                            var cell = cells[count];
                            var value = GetCellValue(ref cell, ref sharedStringTable);
                            columnNumber[count] = GetColumnIndex(cell.CellReference);
                            columnName[count] = GetExcelColumnName(count + 1);
                            properties.Add(columnName[count], value);
                        }
                        listObjects.Add(properties);
                        break;
                    }
                }
            }
            while (reader.Read())
            {
                if (reader.ElementType == typeof(Row))
                {
                    var row = (Row)reader.LoadCurrentElement();
                    var cells = row.Elements<Cell>().ToArray();
                    var cellCount = cells.Length;
                    var properties = new ExpandoObject() as IDictionary<string, Object>;

                    for (var count = 0; count < cellCount; count++)
                    {
                        var cell = cells[count];
                        var value = GetCellValue(ref cell, ref sharedStringTable);
                        var columnIndex = GetColumnIndex(cell.CellReference);
                        var columnFound = false;
                        for (var countColumn = 0; countColumn < columnCount; countColumn++)
                        {
                            if (columnIndex == columnNumber[countColumn])
                            {
                                columnIndex = columnNumber[countColumn];
                                columnFound = true;
                                break;
                            }
                        }

                        if (columnFound)
                        {
                            properties.Add(columnName[columnIndex], value);
                        }
                    }

                    listObjects.Add(properties);
                }
            }

            return listObjects;
        }

        private static int GetColumnIndex(string cellReference)
        {
            //remove digits
            string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);

            int columnNumber = -1;
            int mulitplier = 1;

            //working from the end of the letters take the ASCII code less 64 (so A = 1, B =2...etc)
            //then multiply that number by our multiplier (which starts at 1)
            //multiply our multiplier by 26 as there are 26 letters
            foreach (char c in columnReference.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * ((int)c - 64);

                mulitplier *= 26;
            }

            //the result is zero based so return columnnumber + 1 for a 1 based answer
            //this will match Excel's COLUMN function
            return columnNumber;
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        private static string GetCellValue(ref Cell cell, ref SharedStringTable sharedStringTable)
        {
            string value = null;

            // If the cell does not exist, return an empty string.
            if (cell == null)
            {
                return "";
            }
            if (cell.InnerText.Length > 0)
            {
                value = cell.InnerText;

                if (cell.CellFormula != null)
                {
                    value = cell.CellValue.Text == null ? cell.CellValue.Text : "";
                }
                else if (cell.DataType != null)
                {
                    switch (cell.DataType.Value)
                    {
                        case CellValues.SharedString:

                            // For shared strings, look up the value in the
                            // shared strings table.
                            if (sharedStringTable != null)
                            {
                                value = sharedStringTable.ElementAt(int.Parse(value)).InnerText;
                            }
                            break;
                        case CellValues.Boolean:
                            value = value == "0" ? "FALSE" : "TRUE";
                            break;
                        case CellValues.Date:
                            value = DateTime.FromOADate(double.Parse(value)).ToShortDateString();
                            break;
                        case CellValues.Error:

                            break;
                        case CellValues.Number:

                            break;
                        case CellValues.String:

                            break;
                    }
                }
            }
            else if (cell.CellValue != null && !String.IsNullOrEmpty(cell.CellValue.Text))
            {
                value = cell.CellValue.Text;
            }
            else if (cell.StyleIndex != null)
            {
                value = "";
            }
            return value;
        }

    }
}
