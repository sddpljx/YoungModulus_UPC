using System;
using System.Collections.Generic;
using System.Data;
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

namespace Young_Modulus.Pages
{
    /// <summary>
    /// Interaction logic for LastCalc.xaml
    /// </summary>
    public partial class LastCalc : UserControl
    {
        public LastCalc()
        {
            InitializeComponent();
            try
            {
                ShowLastFriage();
                ShowLastResult();
                ShowLastData();
            }
            catch
            {

            }
        }

        public void ShowLastResult()
        {
            string exePath = Environment.CurrentDirectory;
            string txtPath = null;
            txtPath = exePath + "\\result.txt";
            StreamReader sr = new StreamReader(txtPath, Encoding.Default);
            String line;
            line = sr.ReadLine();
            this.textBoxShowLastAns.Text = line;
            sr.Close();
        }

        public void ShowLastFriage()
        {
            string exePath = Environment.CurrentDirectory;
            string txtPath = null;
            txtPath = exePath + "\\friage.txt";
            StreamReader sr = new StreamReader(txtPath, Encoding.Default);
            String line;
            line = sr.ReadLine();
            this.textBoxShowLastFriage.Text = line;
            sr.Close();
        }

        public void ShowLastData()
        {
            string exePath = Environment.CurrentDirectory;
            string txtPath = null;
            string xlsxPath = null;
            xlsxPath = exePath + "\\dataSave.xlsx";
            ExcelDataTable.dt = ExcelUtility.ExcelToDataTable(xlsxPath, true);
            dataGridShowLast.ItemsSource = ExcelDataTable.dt.DefaultView;
        }

        public void ExportLastData()
        {
            try
            {
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog()
                {
                    Filter = "Excel(*.xlsx *.xls)|*.xlsx;*.xls"
                };
                var result = saveFileDialog.ShowDialog();
                if (result == true)
                {
                    string path = saveFileDialog.FileName;
                    ExcelUtility.DataTableToExcel(path, ExcelDataTable.dt);
                    this.textBoxLastDataAddress.Text = "导出数据成功！路径：" + path;
                }
            }
            catch
            {
                MessageBox.Show("导出数据失败！", "Warnning!");
            }
        }

        private void buttonExportLastData_Click(object sender, RoutedEventArgs e)
        {
            ExportLastData();
        }

        private void buttonUpdateData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ShowLastFriage();
                ShowLastResult();
                ShowLastData();
            }
            catch
            {

            }
        }
    }

    public class ExcelDataTable
    {
        public static DataTable dt;
    }
}
