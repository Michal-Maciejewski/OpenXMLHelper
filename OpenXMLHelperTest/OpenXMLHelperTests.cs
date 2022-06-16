using CsvHelper;
using CsvHelper.Configuration;
using Mapster;
using NUnit.Framework;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace OpenXMLHelperTest
{
    public class OpenXMLHelperTests
    {
        public MemoryStream MS { get; set; }
        public StreamWriter SW { get; set; }

        [OneTimeSetUp]
        public void Setup()
        {
            MS = new MemoryStream();
            SW = new StreamWriter(MS);
            var config = new CsvConfiguration(CultureInfo.CurrentCulture) { Delimiter = ",", HasHeaderRecord = true };
            using var csvWriter = new CsvWriter(SW, config);
            for (var count = 0; count < 6; count++)
            {
                csvWriter.WriteField("A" + count);
            }
            csvWriter.NextRecord();
            for (var row = 0; row < 100; row++)
            {
                for (var column = 0; column < 6; column++)
                {
                    csvWriter.WriteField(row.ToString() + column.ToString());
                }
                csvWriter.NextRecord();
            }
            csvWriter.Flush();
        }

        [Test]
        public void OpenXMLHelper_ReadCSVFileWithHeader_ReturnsValidObjectList()
        {
            var expectedCount = 100;

            var objectList = OpenXMLHelper.OpenXMLHelper.ReadCSVFile(MS.ToArray(), true);
            Assert.NotNull(objectList);

            var actualCount = objectList.Count;
            Assert.AreEqual(expectedCount, actualCount);
        }

        [Test]
        public void OpenXMLHelper_ReadCSVFileWithoutHeader_ReturnsValidObjectList()
        {
            var expectedCount = 101;

            var objectList = OpenXMLHelper.OpenXMLHelper.ReadCSVFile(MS.ToArray(), false);
            Assert.NotNull(objectList);

            var actualCount = objectList.Count;
            Assert.AreEqual(expectedCount, actualCount);
        }

        [Test]
        public void OpenXMLHelper_CreateExcelFile()
        {
            var header = true;
            var list = new List<object> {
                new
                {
                    Name = "Micheal",
                    Address = "Opals Street",
                    Number = 0
                },
                new
                {
                    Name = "Ariel",
                    Address = "Somewhere Street",
                    Number = 1
                },
                new
                {
                    Name = "John",
                    Address = "Triset Street",
                    Number = 2
                },
            };
            OpenXMLHelper.OpenXMLHelper.CreateExcelWorkbookFileToByteArray(list, header);
        }



        [OneTimeTearDown]
        public void KillStreams()
        {
            SW.Close();
            MS.Close();

            SW.Dispose();
            MS.Dispose();
        }
    }
}