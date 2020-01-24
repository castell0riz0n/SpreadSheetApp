using Microsoft.Win32;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using WindowsExcelApp.Classes;
using WindowsExcelApp.Helpers;
using static WindowsExcelApp.Helpers.Helpers;

namespace WindowsExcelApp.Views
{
    /// <summary>
    /// Interaction logic for FindResFromSusThatNotInOrg.xaml
    /// </summary>
    public partial class FindResFromSusThatNotInOrg : Window
    {
        // === Variables area ===
        private static readonly string _fileExtensions = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
        private List<ResourcesDto> _orgDataList = new List<ResourcesDto>();
        private List<ResourcesDto> _susDataList = new List<ResourcesDto>();
        // === === === === === === 

        public FindResFromSusThatNotInOrg()
        {
            InitializeComponent();
            InitializeComponent();
            ButtonActions(false);
        }


        private void Org_File_Open(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = _fileExtensions;
            if (openFileDialog.ShowDialog() == true)
            {
                var fileName = openFileDialog.FileName;
                var excelReaderResult = ConvertExcelToDataSet(fileName);
                if (excelReaderResult.Count > 0)
                {
                    OrgText.Text =
    $"You have selected {openFileDialog.SafeFileName} and we found {excelReaderResult.Count} rows in the file";
                    _orgDataList = excelReaderResult;
                    SusBtn.IsEnabled = true;
                }
                else
                {
                    MessageBox.Show("متاسفانه فایلتون یه مشکلی داره و نمیتونیم اطلاعاتش رو بخونیم", "خطا در وکردیم");
                }
            }
        }

        private void Sus_File_Open(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = _fileExtensions;
            if (openFileDialog.ShowDialog() == true)
            {
                var fileName = openFileDialog.FileName;

                var excelReaderResult = ConvertExcelToDataSet(fileName);

                if (excelReaderResult.Count > 0)
                {
                    SusText.Text =
                        $"You have selected {openFileDialog.SafeFileName} and we found {excelReaderResult.Count} rows in the file";
                    _susDataList = excelReaderResult;
                    StartOp.IsEnabled = true;
                }
                else
                {
                    MessageBox.Show("متاسفانه فایلتون یه مشکلی داره و نمیتونیم اطلاعاتش رو بخونیم", "خطا در وکردیم");
                }
            }
        }

        private void Start_Op(object sender, RoutedEventArgs e)
        {
            var listOfResourcesFound = new List<ResourcesDto>();
            if (_orgDataList.Count > 0 && _susDataList.Count > 0)
            {
                foreach (var res in _susDataList)
                {
                    if (_orgDataList.FirstOrDefault(a => a.Name == res.Name) != null)
                    {
                        continue;
                    }
                    listOfResourcesFound.Add(res);
                }

                string message = $"{listOfResourcesFound.Count} پیدا شدن که تو فایل اصلی هستن  ولی تو فایل تغییر کرده نیستن \n" +
                                 $"الان میگی چیکار کنم ؟ \n" +
                                 $"برای ذخیره و مرج Yes رو بزن \n" +
                                 $"برای ذخیره فقط اونایی که پیدا شده No رو بزن \n" +
                                 $"اگه مرض هم داشتی الکی اومدی تا اینجا Cancel کن \n";
                var userReaction = MessageBox.Show(message, "یه تصمیمی بگیر", MessageBoxButton.YesNoCancel, MessageBoxImage.Question, MessageBoxResult.Yes, MessageBoxOptions.RightAlign);
                _orgDataList.AddRange(listOfResourcesFound);
                _orgDataList = _orgDataList.OrderBy(a => a.Name).ToList();

                var location = SaveDialog();

                switch (userReaction)
                {
                    case MessageBoxResult.Yes:
                        var rows = SaveAsExcel(
                            ConvertJsonToDataTable(listOfResourcesFound.ToJson()),
                            location + "-MissingRows").Result;

                        var newSusResult = SaveAsExcel(
                            ConvertJsonToDataTable(_orgDataList.ToJson()),
                            location + "-Merged").Result;
                        OpTextBox.Text = $"Missing rows Saved : {rows} \n" +
                                         $"Missing rows Added to Org file and Merged file saved : {newSusResult} \n" +
                                         $"Done";
                        ButtonActions(false);
                        break;
                    case MessageBoxResult.No:
                        var missedRows = SaveAsExcel(
                            ConvertJsonToDataTable(listOfResourcesFound.ToJson()),
                            location + "-MissingRows").Result;

                        OpTextBox.Text = $"Missing rows Saved : {missedRows} \n" +
                                         $"Done";
                        ButtonActions(false);
                        break;
                    case MessageBoxResult.Cancel:
                        this.Close();
                        break;
                }
            }
        }

        private string SaveDialog()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = @"C:\";
            saveFileDialog.Title = "حالا یجا سیوش کن بی صاحابو";
            if (saveFileDialog.ShowDialog() == true)
            {
                return saveFileDialog.FileName;
            }

            return @"C:\Updated";
        }

        private void ButtonActions(bool state)
        {
            StartOp.IsEnabled = state;
            SusBtn.IsEnabled = state;
        }

        private void Cancel_Btn(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
