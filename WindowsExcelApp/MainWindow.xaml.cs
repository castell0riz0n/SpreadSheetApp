using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WindowsExcelApp.Views;
using Microsoft.Win32;

namespace WindowsExcelApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Org_To_Sus_Btn_Click(object sender, RoutedEventArgs e)
        {
            FindResFromOrgThatNotInSus window = new FindResFromOrgThatNotInSus();
            window.Show();
        }        
        
        private void Sus_To_Org_Btn_Click(object sender, RoutedEventArgs e)
        {
            FindResFromSusThatNotInOrg window = new FindResFromSusThatNotInOrg();
            window.Show();
        }

        private void Find_Missing_Tr_Btn_Click(object sender, RoutedEventArgs e)
        {
            FindMissingTranslatedRowsInFile window = new FindMissingTranslatedRowsInFile();
            window.Show();
        }


    }

}
