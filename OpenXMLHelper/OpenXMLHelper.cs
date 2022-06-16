namespace OpenXMLHelper
{
    public class OpenXMLHelper
    {
        public static List<object> ReadCSVFile(string filename, bool header)
        {
            using var reader = new StreamReader(filename);
            var objectList = OpenXMLHelperReader.GetListCSV(reader, header);
            return objectList;
        }

        public static List<object> ReadCSVFile(Stream stream, bool header)
        {
            using var reader = new StreamReader(stream);
            var objectList = OpenXMLHelperReader.GetListCSV(reader, header);
            return objectList;
        }

        public static List<object> ReadCSVFile(byte[] byteArray, bool header)
        {
            using var stream = new MemoryStream(byteArray, false);
            using var reader = new StreamReader(stream);
            var objectList = OpenXMLHelperReader.GetListCSV(reader, header);
            return objectList;
        }

        public static List<object> ReadExcelWorkbook(Stream stream, bool header)
        {
            var objectList = OpenXMLHelperReader.GetListExcelWorkbook(stream, header);
            return objectList;
        }

        public static List<object> ReadExcelWorkbook(byte[] byteArray, bool header)
        {
            using Stream stream = new MemoryStream(byteArray, false);
            var objectList = OpenXMLHelperReader.GetListExcelWorkbook(stream, header);
            return objectList;
        }

        public static List<object> ReadExcelWorkbook(string filename, bool header)
        {
            using Stream stream = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var objectList = OpenXMLHelperReader.GetListExcelWorkbook(stream, header);
            return objectList;
        }

        public static void CreateCSVFile<T>(string fileName, T[] record, bool autoMap)
        {
            OpenXMLHelperWriter.CreateCSVFile(fileName, record, autoMap);
        }

        public static void CreateCSVFile<T>(string fileName, List<T> record, bool autoMap)
        {
            OpenXMLHelperWriter.CreateCSVFile(fileName, record, autoMap);
        }

        public static byte[] CreateCSVFileByteArray<T>(T[] record, bool autoMap)
        {
            return OpenXMLHelperWriter.CreateCSVFileToByteArray(record, autoMap);
        }

        public static byte[] CreateCSVFileByteArray<T>(List<T> record, bool autoMap)
        {
            return OpenXMLHelperWriter.CreateCSVFileToByteArray(record, autoMap);
        }

        public static void CreateExcelWorkbookFile<T>(List<T> record, string fileName, bool header)
        {
            OpenXMLHelperWriter.CreateWorkbookUsingOpenXMLWriter(record, fileName, header);
        }

        public static void CreateExcelWorkbookFile<T>(T[] record, string fileName, bool header)
        {
            OpenXMLHelperWriter.CreateWorkbookUsingOpenXMLWriter(record, fileName, header);
        }

        public static byte[] CreateExcelWorkbookFileToByteArray<T>(List<T> record, bool header)
        {
            return OpenXMLHelperWriter.CreateWorkbookUsingOpenXMLWriterToByteArray(record, header);
        }

        public static byte[] CreateExcelWorkbookFileToByteArray<T>(T[] record, bool header)
        {
            return OpenXMLHelperWriter.CreateWorkbookUsingOpenXMLWriterToByteArray(record, header);
        }
    }
}
