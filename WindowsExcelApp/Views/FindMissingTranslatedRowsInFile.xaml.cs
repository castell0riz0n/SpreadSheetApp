using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Windows;
using WindowsExcelApp.Classes;
using WindowsExcelApp.Helpers;
using static WindowsExcelApp.Helpers.Helpers;

namespace WindowsExcelApp.Views
{
    /// <summary>
    /// Interaction logic for FindMissingTranslatedRowsInFile.xaml
    /// </summary>
    public partial class FindMissingTranslatedRowsInFile : Window
    {
        // === Variables area ===
        private static readonly string _fileExtensions = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
        private List<ResourcesDto> _orgDataList = new List<ResourcesDto>();
        // === === === === === === 
        public FindMissingTranslatedRowsInFile()
        {
            InitializeComponent();
        }

        private void Cancel_Btn(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Org_File_Open(object sender, RoutedEventArgs e)
        {
            OpTextBox.Text = String.Empty;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = _fileExtensions;
            if (openFileDialog.ShowDialog() == true)
            {
                var fileName = openFileDialog.FileName;
                var excelReaderResult = ConvertExcelToDataSet(fileName);
                if (excelReaderResult.Count > 0)
                {
                    OpTextBox.Text =
                        $"You have selected {openFileDialog.SafeFileName} and we found {excelReaderResult.Count} rows in the file \n";
                    _orgDataList = excelReaderResult;
                    Find_Missing_Translated_Rows();
                }
                else
                {
                    MessageBox.Show("متاسفانه فایلتون یه مشکلی داره و نمیتونیم اطلاعاتش رو بخونیم", "خطا در وکردیم");
                }
            }
        }

        private void Find_Missing_Translated_Rows()
        {
            var notTranslated = new List<ResourcesDto>();
            foreach (var row in _orgDataList)
            {
                if (string.IsNullOrEmpty(row.Value))
                {
                    notTranslated.Add(row);
                    OpTextBox.Text = OpTextBox.Text + row.Name + "\n";
                }
                else if (Utilities.WordIsInPersianOrArabic(row.Value))
                {
                    notTranslated.Add(row);
                    OpTextBox.Text = OpTextBox.Text + row.Name + "\n";
                }
            }

            if (notTranslated.Count > 0)
            {
                string message = $"{notTranslated.Count} found, What should I do now ?\n" +
                                 $"(OK) To Save as a File\n" +
                                 $"(Cancel) See result in Final Result box \n";

                // wait for user to react
                var userReaction = MessageBox.Show(message, "Make a Decision : ",MessageBoxButton.OKCancel,MessageBoxImage.Question,MessageBoxResult.OK);

                switch (userReaction)
                {
                    case MessageBoxResult.OK:
                        var location = SaveDialog();
                        var rows = SaveAsExcel(
                            ConvertJsonToDataTable(notTranslated.ToJson()),
                            location + "-MissingTranslated").Result;
                        OpTextBox.Text = OpTextBox.Text + $"File saved in {rows} \n";
                        break;
                    case MessageBoxResult.Cancel:
                        break;
                }
            }
            else
            {
                MessageBox.Show("Hayya!!! Everything are up and Ok", "Eyval", MessageBoxButton.OK,
                    MessageBoxImage.Information, MessageBoxResult.OK);
            }

        }

        private string SaveDialog()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "حالا یجا سیوش کن بی صاحابو";
            saveFileDialog.InitialDirectory = @"C:\";
            if (saveFileDialog.ShowDialog() == true)
            {
                return saveFileDialog.FileName;
            }

            return @"C:\Updated";
        }
    }
}
