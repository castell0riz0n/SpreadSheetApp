using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using ExcelApp.Classes;
using ExcelApp.Utilities;
using Newtonsoft.Json;

namespace ExcelApp
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            await Menu();
        }

        public static async Task Menu()
        {
            try
            {
                switch (SelectMenu())
                {
                    case 1:
                        await StartOperation();
                        await Menu();
                        break;
                    case 2:
                        await FindRowsNotInOtherFiles();
                        await Menu();
                        break;
                    case 3:
                        // await Parto.MapProperies();
                        await Menu();
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                await Menu();
            }

        }

        public static int SelectMenu()
        {
            try
            {
                Console.WriteLine("Start Map and Merge (1): ");
                Console.WriteLine("Check For Diff and Add to Other (2) : ");
                Console.Write("Please Type an option number To start : ");
                return Int32.Parse(Console.ReadLine() ?? throw new InvalidOperationException("Please just enter one the available options number"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine("press any key to see menu . . .");
                Console.ReadLine();
                return 0;
            }

        }


        private static List<string> FindFiles()
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

        private static async Task<ResourcesBaseDto> ReadFile(string fileAddress)
        {
            try
            {
                var data = new ResourcesBaseDto();
                using (StreamReader reader = new StreamReader(fileAddress))
                {
                    var json = await reader.ReadToEndAsync();
                    data = json.ToObjectFromJson<ResourcesBaseDto>();
                }

                if (data?.Sheet1.Count == 0)
                {
                    Console.WriteLine("cannot deserialize json");
                    await Program.Menu();
                }

                Console.WriteLine($"{data.Sheet1.Count} Resource Name Found");
                await Task.Delay(1000);
                return data;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return new ResourcesBaseDto();
            }
        }

        private static async Task<List<ResourcesDto>> ReadFileByAddress(string fileAddress)
        {
            try
            {
                var data = new List<ResourcesDto>();
                using (StreamReader reader = new StreamReader(fileAddress))
                {
                    var json = await reader.ReadToEndAsync();
                    data = json.ToObjectFromJson<List<ResourcesDto>>();
                }

                if (data?.Count == 0)
                {
                    Console.WriteLine("cannot deserialize json");
                    await Program.Menu();
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

        private static async Task StartOperation()
        {
            try
            {
                var files = FindFiles();
                if (files.Count == 0)
                {
                    Console.WriteLine(@"No files found in C:\excels");
                    await Program.Menu();
                }

                foreach (var file in files)
                {
                    var fileName = file.Split('\\', StringSplitOptions.RemoveEmptyEntries);
                    var currentFileData = await ReadFile(file);
                    if (currentFileData.Sheet1.Count == 0)
                    {
                        Console.WriteLine("No Data returned from ReadFile Method");
                        await Program.Menu();
                    }

                    var mapped = await MapData(currentFileData.Sheet1, currentFileData.Sheet2);

                    var addIsNotInList = await AddMissingNames(currentFileData.Sheet1, mapped);

                    //using (StreamWriter writer = new StreamWriter(@"c:\excels\updatedData.json", true))
                    //{
                    //    writer.WriteLine(addIsNotInList.ToJson());
                    //}

                    DataTable dt = (DataTable)JsonConvert.DeserializeObject(addIsNotInList.ToJson(), (typeof(DataTable)));

                    await ConvertJsonToExcel(dt, $"MergedData-{new Random().Next(10, 100)}");

                    Console.Read();
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                await Program.Menu();
            }
        }

        private static async Task<List<ResourcesDto>> MapData(List<ResourcesDto> sheet1, List<ResourcesDto> sheet2)
        {

            var updatedList = new List<ResourcesDto>();
            var UpdatedResourceList = updatedList;
            Parallel.ForEach(sheet2, rowData =>
            {
                var findedInSheet1 = sheet1.FirstOrDefault(a => a.Name == rowData.Name);

                if (string.IsNullOrEmpty(rowData.Value))
                {
                    if (!string.IsNullOrEmpty(findedInSheet1?.Value))
                    {
                        rowData.Value = findedInSheet1.Value;
                    }
                    UpdatedResourceList.Add(rowData);
                    return;
                }

                if (Utilities.Utilities.WordIsInPersianOrArabic(rowData.Value))
                {
                    if (findedInSheet1 == null)
                    {
                        UpdatedResourceList.Add(rowData);
                        return;
                    }

                    if (string.IsNullOrEmpty(findedInSheet1.Value))
                    {
                        UpdatedResourceList.Add(rowData);
                        return;
                    }
                    if (Utilities.Utilities.WordIsInPersianOrArabic(findedInSheet1.Value))
                    {
                        UpdatedResourceList.Add(rowData);
                        return;
                    }

                    rowData.Value = findedInSheet1.Value;
                    UpdatedResourceList.Add(rowData);
                    return;
                }
                UpdatedResourceList.Add(rowData);
            });
            updatedList = updatedList.OrderBy(a => a.Name).ToList();
            return updatedList;
        }

        private static async Task<List<ResourcesDto>> AddMissingNames(List<ResourcesDto> previouslyMappedList,
            List<ResourcesDto> bigList)
        {
            var newMappedList = previouslyMappedList;
            var origList = bigList;
            var list = newMappedList;
            //Parallel.ForEach(origList, a =>
            //{

            //});

            foreach (var res in origList)
            {
                var findedInPreviouslyMappedList = previouslyMappedList.FirstOrDefault(f => f.Name == res.Name);
                if (findedInPreviouslyMappedList == null)
                {
                    list.Add(res);
                }
            }


            newMappedList = list.OrderBy(d => d.Name).ToList();

            return newMappedList;

        }

        private static async Task<bool> FindMissingTranslated()
        {
            return false;
        }


        private static async Task ConvertJsonToExcel(DataTable jsonDt, string fileName)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(jsonDt ,"Sample Sheet");
                workbook.SaveAs($"C:\\excels\\{fileName}.xlsx");
            }

        }

        private static async Task FindRowsNotInOtherFiles()
        {
            var originalFile = await ReadFileByAddress(@"c:\excels\originalFile.json");
            var comparedFile = await ReadFileByAddress(@"c:\excels\comparedFile.json");

            var orgData = originalFile;
            var compData = comparedFile;

            var resourcesNotInCompFile = new List<ResourcesDto>();

            foreach (var data in orgData)
            {
                var founded = compData.FirstOrDefault(a => a.Name == data.Name);

                if (founded != null)
                {
                    continue;
                }
                resourcesNotInCompFile.Add(data);
            }

            if (resourcesNotInCompFile.Count == 0)
            {
                Console.WriteLine("No Resources found to add");
            }
            compData.AddRange(resourcesNotInCompFile);

            compData = compData.DistinctBy(a => a.Name).ToList();

            compData = compData.OrderBy(a => a.Name).ToList();
            DataTable dt = (DataTable)JsonConvert.DeserializeObject(compData.ToJson(), (typeof(DataTable)));
            string fileName = $"UpdatedExcel-{new Random().Next(1000, 5000)}";
            await ConvertJsonToExcel(dt, fileName);

            Console.WriteLine($"File Generated Successfully : {fileName}");
        }

        private static async Task MergeFilesTogether()
        {
            string baseFile = @"";
        }
    }
}
