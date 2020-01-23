using System.Collections.Generic;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using ExcelDataReader;
using Newtonsoft.Json;

namespace ExcelApp.Classes
{
    public static class ConvertExcelToJson
    {
        private static readonly List<string> FilesAddress = new List<string>
        {
            @"c:\excels\OriginalData.xlsx",
            @"c:\excels\UpdatedData.xlsx",

        };
        public static void GenerateJson()
        {
            
        }

        private static void UsingExcelDataReader(string excelFile, string jsonFile)
        {
            //string inputFile = @"c:\excels\OriginalData.xlsx";
            using (var inFile = File.Open(excelFile, FileMode.Open, FileAccess.Read))
            using (var outFile = File.CreateText(jsonFile))
            using (var reader = ExcelReaderFactory.CreateReader(inFile,
                new ExcelReaderConfiguration { FallbackEncoding = Encoding.GetEncoding(1252) }))
            using (var writer = new JsonTextWriter(outFile))
            {
                writer.Formatting = Formatting.Indented;
                writer.WriteStartArray();
                //SKIP FIRST ROW, it's TITLES.
                reader.Read();
                do
                {
                    while (reader.Read())
                    {
                        //We don't need empty object
                        var Name = reader.GetString(0);
                        if (string.IsNullOrEmpty(Name)) break;

                        writer.WriteStartObject();
                        //Select Columns and values
                        writer.WritePropertyName("Name");
                        writer.WriteValue(Name);

                        writer.WritePropertyName("Value");
                        writer.WriteValue(reader.GetString(1));

                        writer.WritePropertyName("Comment");
                        writer.WriteValue(reader.GetString(2));

                        writer.WriteEndObject();
                    }
                } while (reader.NextResult());

                writer.WriteEndArray();
            }
        }
    }
}