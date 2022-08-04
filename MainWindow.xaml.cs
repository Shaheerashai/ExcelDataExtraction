using System;
using System.Collections.Generic;
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

namespace ExcelDataExtraction
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

        private void Start_Button_Click(object sender, RoutedEventArgs e)
        {
            string sInputFile = FileDirTextBox.Text;
            MainFunctionStart(sInputFile);
        }

        private void Browse_Button_Click(object sender, RoutedEventArgs e)
        {
            var SelectedFile = new Microsoft.Win32.OpenFileDialog() { Filter = "Excel Files|*.xls;*.xlsx;*.xlsm" };
            var result = SelectedFile.ShowDialog();
            if (result == false) return;
            FileDirTextBox.Text = SelectedFile.FileName;
        }
        private void MainFunctionStart(string sInputFile)
        {
            ExcelHelper ex = new ExcelHelper(sInputFile);
            Dictionary<string, List<string>> ExcelDataDict = ex.ExtractExcelData(sInputFile, "Read");
            ex.GenerateOutputText(ExcelDataDict);
            MessageBox.Show("Process Finished");
            this.Close();
        }
    }
}
