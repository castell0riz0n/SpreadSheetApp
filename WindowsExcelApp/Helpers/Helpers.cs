using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WindowsExcelApp.Classes;
using ClosedXML.Excel;
using ExcelDataReader;
using Newtonsoft.Json;

namespace WindowsExcelApp.Helpers
{
    public static class Helpers
    {
        public static async Task<string> SaveAsExcel(DataTable jsonDt, string fileName)
        {
            using (var workbook = new XLWorkbook())
            {
                try
                {
                    DateTime opDate = DateTime.Now;
                    fileName = $"{fileName}-H{opDate.Hour}M{opDate.Minute}S{opDate.Second}";
                    var worksheet = workbook.Worksheets.Add(jsonDt, "Data");
                    workbook.SaveAs($"{fileName}.xlsx");
                    return fileName;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    throw;
                }

            }

        }

        public static async Task<List<ResourcesDto>> ReadJsonFileByName(string fileName)
        {
            try
            {
                var data = new List<ResourcesDto>();
                using (StreamReader reader = new StreamReader($"D:\\ResourceData\\Json\\{fileName}.json"))
                {
                    var json = await reader.ReadToEndAsync();
                    data = json.ToObjectFromJson<List<ResourcesDto>>();
                }

                if (data?.Count == 0)
                {
                    Console.WriteLine("cannot deserialize json");
                    //await Program.Menu();
                }

                Console.WriteLine($"{data.Count} Resource Name Found");
                await Task.Delay(1000);
                return data;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return new List<ResourcesDto>();
            }
        }

        public static List<string> FindFiles()
        {
            try
            {
                var filesPath = Directory.GetFiles(@"c:\excels", "*.json", SearchOption.AllDirectories);

                return filesPath.ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine("cannot find files in given directory");
                return new List<string>();
            }
        }

        public static void ConvertExcelToJson(string fileName)
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                var inFilePath = $"D:\\ResourceData\\Excel\\{fileName}.xlsx";
                var outFilePath = $"D:\\ResourceData\\Json\\{fileName}.json";

                using (var inFile = File.Open(inFilePath, FileMode.Open, FileAccess.Read))
                using (var outFile = File.CreateText(outFilePath))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(inFile, new ExcelReaderConfiguration()
                    { FallbackEncoding = Encoding.GetEncoding(1252) }))
                    using (var writer = new JsonTextWriter(outFile))
                    {
                        writer.Formatting = Formatting.Indented; //I likes it tidy
                        writer.WriteStartArray();
                        //reader.Read(); //SKIP FIRST ROW, it's TITLES.
                        do
                        {
                            while (reader.Read())
                            {
                                //peek ahead? Bail before we start anything so we don't get an empty object
                                var name = reader.GetString(0);
                                if (string.IsNullOrEmpty(name))
                                {
                                    break;
                                }

                                writer.WriteStartObject();
                                writer.WritePropertyName("Name");
                                writer.WriteValue(name);

                                writer.WritePropertyName("Value");
                                writer.WriteValue(reader.GetString(1));

                                writer.WritePropertyName("Comment");
                                writer.WriteValue(reader.GetString(2));
                                /*
                                <iframe src="https://channel9.msdn.com/Shows/Azure-Friday/Erich-Gamma-introduces-us-to-Visual-Studio-Online-integrated-with-the-Windows-Azure-Portal-Part-1/player" width="960" height="540" allowFullScreen frameBorder="0"></iframe>
                                 */

                                writer.WriteEndObject();
                            }
                        } while (reader.NextResult());
                        writer.WriteEndArray();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        public static List<ResourcesDto> ConvertExcelToDataSet(string filePath)
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                var list = new List<ResourcesDto>();
                //var outFilePath = $"D:\\ResourceData\\Json\\{fileName}.json";

                using (var inFile = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(inFile))
                    {


                        // 2. Use the AsDataSet extension method
                        var result = reader.AsDataSet();

                        // The result of each spreadsheet is in result.Tables
                        //DataTable resources = result.Tables["Data"];

                        list = result.Tables["Data"].AsEnumerable()
                            .Select(row => new ResourcesDto
                            {
                                Name = row[0].ToString(),
                                Value = row[1].ToString(),
                                Comment = row[2].ToString()
                            }).ToList();
                        //var list = result.Tables["Data"].ToList<ResourcesDto>();
                    }
                }

                return list;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        public static void DeleteExistingFilesInDirectory(string path)
        {
            try
            {
                DirectoryInfo directory = new DirectoryInfo(path);
                foreach (FileInfo file in directory.GetFiles())
                {
                    file.Delete();
                }

                foreach (DirectoryInfo dir in directory.GetDirectories())
                {
                    dir.Delete(true);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        public static DataTable ConvertJsonToDataTable(string json)
        {
            return (DataTable)JsonConvert.DeserializeObject(json, (typeof(DataTable)));
        }

        public static string GetInput(string Prompt)
        {
            string Result = "";
            do
            {
                Console.Write(Prompt + ": ");
                Result = Console.ReadLine();
                if (string.IsNullOrEmpty(Result))
                {
                    Console.WriteLine("Empty input, please try again");
                }
            } while (string.IsNullOrEmpty(Result));
            return Result;
        }

        public static void OpenResultFolder()
        {
            Process.Start("explorer.exe", @"D:\ResourceData\Result");
        }
    }
}