using ExcelApp.Classes;
using ExcelApp.Utilities;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using static ExcelApp.Helpers.Helpers;
using static ExcelApp.Utilities.Utilities;

namespace ExcelApp
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.ResetColor();
            Console.WriteLine("Hello Parnaz(KOPOLI) ! Hope you doing Great as usual");
            await Menu();
        }

        public static async Task Menu()
        {
            try
            {
                switch (SelectMenu())
                {
                    case "1":
                        await FindResFromOrgThatNotInSus();
                        await Menu();
                        break;
                    case "2":
                        await FindResFromSusThatNotInOrg();
                        await Menu();
                        break;
                    case "3":
                        // await Parto.MapProperies();
                        await Menu();
                        break;
                }

                await Menu();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                await Menu();
            }

        }

        public static string SelectMenu()
        {
            try
            {
                Console.WriteLine("(1) Find rows from ORG file that not in Sus file : ");
                Console.WriteLine("(2) Find rows from Sus file that not in Org file : ");
                return GetInput("Please Type an option number To start");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine("press any key to see menu . . .");
                Console.ReadLine();
                Console.Clear();
                return "0";
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
                    var currentFileData = await ReadJsonFile(file);
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

                    await SaveAsExcel(dt, $"MergedData");

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

                if (WordIsInPersianOrArabic(rowData.Value))
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
                    if (WordIsInPersianOrArabic(findedInSheet1.Value))
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

        private static async Task FindRowsNotInOtherFiles()
        {
            var originalFile = await ReadJsonFileByName("originalFile");
            var comparedFile = await ReadJsonFileByName("comparedFile");

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
            await SaveAsExcel(dt, fileName);

            Console.WriteLine($"File Generated Successfully : {fileName}");
        }

        private static async Task FindResFromOrgThatNotInSus()
        {
            try
            {
                //0. Variables and warnings area
                var listOfResourcesFound = new List<ResourcesDto>();
                var userReaction = "";

                Console.WriteLine(@"Please put your files in C:\ResourceData\Excel\");
                //1. Convert Org and Sus excels to json
                var org = ConvertExcelToDataSet("org");
                var sus = ConvertExcelToDataSet("sus");

                //2. read Org and Sus json files 
                //var org = await Helpers.Helpers.ReadJsonFileByName("org");
                //var sus = await Helpers.Helpers.ReadJsonFileByName("sus");

                //3. find all res that is in Org but not in Sus
                foreach (var res in org)
                {
                    if (sus.FirstOrDefault(a => a.Name == res.Name) != null)
                    {
                        continue;
                    }
                    listOfResourcesFound.Add(res);
                }

                //4. print number of result
                if (listOfResourcesFound.Count == 0)
                {
                    Console.WriteLine("Hayyaa!!! files are all up . No resources found to add to the Sus file :-D");
                    await ReturnToMenu();
                }

                //5. ask merge with sus
                Console.Write($"{listOfResourcesFound.Count} line of resources found that they are not in Sus file :-( \n What should we do? \n\t To Save in a separate file Press (s) \n\t To add missing lines and merge with Sus Press (u)");

                userReaction = GetInput("(s/u)?");

                //6. base on answer | if no just save res | if yes merge and then save the res
                switch (userReaction.ToLower())
                {
                    case "s":
                        var saveResult = await SaveAsExcel(
                            ConvertJsonToDataTable(listOfResourcesFound.ToJson()),
                            "MissingResourcesRow");
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Save Successfully as {saveResult}.xlsx");
                        Console.ResetColor();
                        OpenResultFolder();
                        await ReturnToMenu();
                        break;
                    case "u":
                        sus.AddRange(listOfResourcesFound);
                        sus = sus.OrderBy(a => a.Name).ToList();
                        var rows = await SaveAsExcel(
                            ConvertJsonToDataTable(listOfResourcesFound.ToJson()),
                            "MissingResourcesRow");
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Save Successfully as {rows}.xlsx");

                        var newSusResult = await SaveAsExcel(
                            ConvertJsonToDataTable(sus.ToJson()),
                            "Updated-Sus");
                        Console.ResetColor();
                        OpenResultFolder();
                        await ReturnToMenu();
                        break;
                }

                Console.WriteLine("You Entered wrong answer");
                await ReturnToMenu();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                await Menu();
            }
        }

        private static async Task FindResFromSusThatNotInOrg()
        {
            try
            {
                //0. Variables and warnings area
                var listOfResourcesFound = new List<ResourcesDto>();
                var userReaction = "";

                Console.WriteLine(@"Please put your files in C:\ResourceData\Excel\");
                //1. Convert Org and Sus excels to Dataset
                var org = ConvertExcelToDataSet("org");
                var sus = ConvertExcelToDataSet("sus");

                //2. read Org and Sus json files 
                //var org = await Helpers.Helpers.ReadJsonFileByName("org");
                //var sus = await Helpers.Helpers.ReadJsonFileByName("sus");

                //3. find all res that is in Org but not in Sus
                foreach (var res in sus)
                {
                    if (org.FirstOrDefault(a => a.Name == res.Name) != null)
                    {
                        continue;
                    }
                    listOfResourcesFound.Add(res);
                }

                //4. print number of result
                if (listOfResourcesFound.Count == 0)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Hayyaa!!! files are all up . no differences found :-D");
                    Console.ResetColor();
                    await ReturnToMenu();
                }

                //5. ask merge with sus
                Console.Write($"{listOfResourcesFound.Count} line of resources found that they are not in ORG file :-( \n What should we do? \n\t To Save in a separate file Press (s) \n\t To add missing lines and merge with Org Press (u)");

                userReaction = GetInput("(s/u)?");

                //6. base on answer | if no just save res | if yes merge and then save the res
                switch (userReaction.ToLower())
                {
                    case "s":
                        var saveResult = await SaveAsExcel(
                            ConvertJsonToDataTable(listOfResourcesFound.ToJson()),
                            "MissingResourcesRow");
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Successfully saved as {saveResult}.xlsx");
                        Console.ResetColor();
                        OpenResultFolder();
                        await ReturnToMenu();
                        break;
                    case "u":
                        org.AddRange(listOfResourcesFound);
                        org = org.OrderBy(a => a.Name).ToList();

                        var rows = await SaveAsExcel(
                            ConvertJsonToDataTable(listOfResourcesFound.ToJson()),
                            "MissingResourcesRow");
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Successfully saved as {rows}.xlsx");

                        var newSusResult = await SaveAsExcel(
                            ConvertJsonToDataTable(org.ToJson()),
                            "Updated-Org");

                        Console.WriteLine($"Successfully saved as {newSusResult}.xlsx");
                        Console.ResetColor();
                        OpenResultFolder();
                        await ReturnToMenu();
                        break;
                }

                Console.WriteLine("You Entered wrong answer. lets back to menu");
                await ReturnToMenu();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                await ReturnToMenu();
            }
        }

        private static async Task FindAndReplaceMissingTranslated()
        {
            //0. Variable Area
            
            int translated = 0;
            int hasNoTranslation = 0;
            var updatedList = new List<ResourcesDto>();
            var hasNoTranslationList = new List<ResourcesDto>();

            //1. get files
            var org = ConvertExcelToDataSet("org");
            var sus = ConvertExcelToDataSet("sus");

            //2. check for missed translations

            Parallel.ForEach(org, rowData =>
            {
                var findedInSus = sus.FirstOrDefault(a => a.Name == rowData.Name);

                if (string.IsNullOrEmpty(rowData.Value))
                {
                    if (!string.IsNullOrEmpty(findedInSus?.Value))
                    {
                        rowData.Value = findedInSus.Value;
                        translated = translated + 1;
                    }
                    updatedList.Add(rowData);
                    return;
                }

                if (WordIsInPersianOrArabic(rowData.Value))
                {
                    if (findedInSus == null)
                    {
                        updatedList.Add(rowData);
                        hasNoTranslationList.Add(rowData);

                        hasNoTranslation = hasNoTranslation + 1;
                        return;
                    }

                    if (string.IsNullOrEmpty(findedInSus.Value))
                    {
                        updatedList.Add(rowData);
                        hasNoTranslationList.Add(rowData);
                        hasNoTranslation = hasNoTranslation + 1;
                        return;
                    }
                    if (WordIsInPersianOrArabic(findedInSus.Value))
                    {
                        updatedList.Add(rowData);
                        hasNoTranslationList.Add(rowData);
                        hasNoTranslation = hasNoTranslation + 1;
                        return;
                    }

                    rowData.Value = findedInSus.Value;
                    translated = translated + 1;
                    updatedList.Add(rowData);
                    return;
                }
                updatedList.Add(rowData);
            });

            if (translated>0)
            {
                Console.WriteLine($"{translated} of rows translated :-D");
                var newSusResult = await SaveAsExcel(
                    ConvertJsonToDataTable(updatedList.ToJson()),
                    "Org-Translated-By-Sus");
            }

            if (hasNoTranslation > 0 & hasNoTranslationList.Count > 0)
            {
                Console.WriteLine($"Also we can not find any translation for about {hasNoTranslation} rows \n please see the generated excel files that they don't have any translation");

                var newSusResult = await SaveAsExcel(
                    ConvertJsonToDataTable(hasNoTranslationList.ToJson()),
                    "No-Translation-Found");
            }

            await ReturnToMenu();
        }
    }
}
