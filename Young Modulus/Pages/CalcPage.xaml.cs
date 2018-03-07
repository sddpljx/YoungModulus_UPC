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
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Reflection;

namespace Young_Modulus.Pages
{
    /// <summary>
    /// Interaction logic for BasicPage1.xaml
    /// </summary>
    public partial class BasicPage1 : UserControl
    {
        public BasicPage1()
        {
            InitializeComponent();
            FirstRun();
        }

        public void ReadExcelData()
        {
            string address = "";
            string filePath = this.textBoxDataAddress.Text;
            int j = 0;
            if (filePath[0] == '"')
            {
                address = filePath.Substring(1, filePath.Length - 2);
            }
            else
            {
                address = filePath;
            }
            ExcelData.dataTable = ExcelUtility.ExcelToDataTable(address, true);
        }

        public void FirstRun()
        {
            ExcelData.dataTable = new DataTable();
            DataRow dR = ExcelData.dataTable.NewRow();
            ExcelData.dataTable.Columns.Add(new DataColumn("序号", typeof(int)));
            ExcelData.dataTable.Columns.Add(new DataColumn("砝码质量(g)", typeof(int)));
            ExcelData.dataTable.Columns.Add(new DataColumn("d1(mm)", typeof(double)));
            ExcelData.dataTable.Columns.Add(new DataColumn("d2(mm)", typeof(double)));
            ExcelData.dataTable.Columns.Add(new DataColumn("d3(mm)", typeof(double)));
            ExcelData.dataTable.Columns.Add(new DataColumn("劈尖长度(mm)", typeof(double)));
            dR["序号"] = 1;
            dR["砝码质量(g)"] = 0;
            dR["d1(mm)"] = 0.796;
            dR["d2(mm)"] = 0.818;
            dR["d3(mm)"] = 0.788;
            dR["劈尖长度(mm)"] = 6.012;
            ExcelData.dataTable.Rows.Add(dR);
            dataGridShow.ItemsSource = ExcelData.dataTable.DefaultView;
        }

        private void buttonIni_Click(object sender, RoutedEventArgs e)
        {
            this.textBoxNum.Text = "60";
            this.sliderNum.Value = 60;
            this.textBoxShowAns.Text = "";
            this.textBoxDataAddress.Text = "C:\\data.xlsx";
            this.textBlockShowCalcStatement.Text = "";
            ExcelData.dataTable = null;
            GloableVar.dataNum = 0;
            GloableVar.whetherReadSuc = 0;
            dataGridShow.ItemsSource = null;
            FirstRun();
            this.textBlockWhetherExport.Text = "";
        }

        private void buttonReadData_Click(object sender, RoutedEventArgs e)
        {
            GloableVar.fringeNum = Convert.ToInt32(this.textBoxNum.Text);
            ReadExcelData();
            if (GloableVar.whetherReadSuc == 1)
            {
                this.textBlockShowCalcStatement.Text = "读取成功！";
                DataGridComb();
            }
            else
            {
                this.textBlockShowCalcStatement.Text = "读取失败！";
            }
            Timer();
        }

        public void Timer()
        {
            GloableVar.dTimer.Tick += DTimer_Tick;
            GloableVar.dTimer.Interval = new TimeSpan(0, 0, 5);
            GloableVar.dTimer.Start();
        }

        private void DTimer_Tick(object sender, EventArgs e)
        {
            GloableVar.dTimer.Stop();
            this.textBlockShowCalcStatement.Text = "";
            this.textBlockWhetherExport.Text = "";
        }

        private void buttonCalc_Click(object sender, RoutedEventArgs e)
        {
            int num = ExcelData.dataTable.Rows.Count;
            if (GloableVar.whetherReadSuc == 1 | num >= 2)
            {
                CalculateYM yM = new CalculateYM();
                GloableVar.fringeNum =Convert.ToInt32(this.textBoxNum.Text);
                double result = yM.CalResult(GloableVar.fringeNum);
                this.textBoxShowAns.Text = Convert.ToString(Math.Round(result, 6));
                string exePath = Environment.CurrentDirectory;
                string txtPath = null;
                txtPath = exePath + "\\result.txt";
                TxtWrite(txtPath, this.textBoxShowAns.Text);
                string txtPath1 = null;
                txtPath1 = exePath + "\\friage.txt";
                TxtWrite(txtPath1, this.textBoxNum.Text);
                string xlsxPath = null;
                xlsxPath = exePath + "\\dataSave.xlsx";
                ExcelUtility.DataTableToExcel(xlsxPath, ExcelData.dataTable);
                ExportCurrentData();
            }
            else
            {
                MessageBox.Show("没有输入数据或者读取数据失败，无法计算！", "Warnning!");
            }
        }

        private void buttonSelectFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog()
            {
                Filter = "Excel(*.xlsx *.xls)|*.xlsx;*.xls"
            };
            var result = openFileDialog.ShowDialog();
            if (result == true)
            {
                string path = openFileDialog.FileName;
                this.textBoxDataAddress.Text = path;
                ExcelData.excelPath = path;
                ReadExcelData();
                DataGridComb();
                this.textBlockShowCalcStatement.Text = "读取成功！";
            }
            else
            {
                this.textBlockShowCalcStatement.Text = "读取失败！";
            }
            Timer();
        }

        public void ExportCurrentData()
        {
            System.DateTime currentTime = new System.DateTime();
            currentTime = System.DateTime.Now;
            string folderPath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString() + "\\Young Modulus\\";
            string excelPath = currentTime.ToString("yyyyMMdd_HHmmss") + "_" + this.textBoxNum.Text + "_" + this.textBoxShowAns.Text;
            ExcelUtility.DataTableToExcel(folderPath + excelPath+".xlsx", ExcelData.dataTable);
        }

        private void buttonExportData_Click(object sender, RoutedEventArgs e)
        {
            ExportData();
        }

        private void buttonExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.Close();
        }

        private void SliderNum_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            int sliderNumber;
            sliderNumber = Convert.ToInt32(this.sliderNum.Value);
            GloableVar.fringeNum = sliderNumber;
            this.textBoxNum.Text = Convert.ToString(sliderNumber);
        }

        public void DataGridComb()
        {
            dataGridShow.ItemsSource = ExcelData.dataTable.DefaultView;
        }

        public void TxtWrite(string txtPath, string con)
        {
            FileStream file = new FileStream(txtPath, FileMode.Create);
            StreamWriter sw = new StreamWriter(file);
            string str = con;
            byte[] data = System.Text.Encoding.Default.GetBytes(str);
            file.Write(data, 0, data.Length);
            file.Flush();
            file.Close();
        }

        public void ExportData()
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
                    ExcelUtility.DataTableToExcel(path, ExcelData.dataTable);
                    this.textBlockWhetherExport.Text = "导出成功！";
                }
                Timer();
            }
            catch
            {
                MessageBox.Show("导出数据失败！", "Warnning!");
            }
        }

        class CalculateYM
        {
            public double CalResult(int tiaoshu)
            {
                DataTable dT = ExcelData.dataTable;
                GloableVar.dataNum = ExcelData.dataTable.Rows.Count;
                int num = GloableVar.dataNum;
                double[,] data = new double[num + 5, 5];
                int j = 0;
                for (int i = 0; i < num * 5; i += 5)
                {
                    if (j == num)
                    {
                        break;
                    }
                    data[j, 0] = Convert.ToInt32(dT.Rows[j][1]);
                    data[j, 1] = Convert.ToDouble(dT.Rows[j][2]);
                    data[j, 2] = Convert.ToDouble(dT.Rows[j][3]);
                    data[j, 3] = Convert.ToDouble(dT.Rows[j][4]);
                    data[j, 4] = Convert.ToDouble(dT.Rows[j][5]);
                    j++;
                }
                double[] d = new double[num * 5];  //dd[i]用于记录直径,ee[i]用于记录杨氏模量
                double[] ee = new double[num * 5];
                double st;
                double result = 0;
                for (int i = 0; i < num; i++)
                {
                    st = (data[i, 1] + data[i, 2] + data[i, 3]) / 3.0;//di的均值
                    d[i] = tiaoshu / st / 2 * 5893 * data[i, 4] * 0.000001;//直径
                }

                for (int i = 1; i < num; i++)
                {
                    ee[i] = 4 * 0.47 * d[0] * data[i, 0] * 0.0098 / 3.1415 / d[i] / d[i] / (d[0] - d[i]);//杨氏模量

                }
                for (int i = 1; i < num; i++)
                {
                    result = result + ee[i];
                }
                result = result / (num - 1);
                return result;
            }
        }

    }

    public class ExcelUtility
    {
        /// <summary>
        /// 将excel导入到datatable
        /// </summary>
        /// <param name="filePath">excel路径</param>
        /// <param name="isColumnName">第一行是否是列名</param>
        /// <returns>返回datatable</returns>
        public static DataTable ExcelToDataTable(string filePath, bool isColumnName)
        {
            DataTable dataTable = null;
            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            int startRow = 0;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet
                        dataTable = new DataTable();
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数
                            if (rowCount > 0)
                            {
                                IRow firstRow = sheet.GetRow(0);//第一行
                                int cellCount = firstRow.LastCellNum;//列数

                                //构建datatable的列
                                if (isColumnName)
                                {
                                    startRow = 1;//如果第一行是列名，则从第二行开始读取
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        cell = firstRow.GetCell(i);
                                        if (cell != null)
                                        {
                                            if (cell.StringCellValue != null)
                                            {
                                                column = new DataColumn(cell.StringCellValue);
                                                dataTable.Columns.Add(column);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        column = new DataColumn("column" + (i + 1));
                                        dataTable.Columns.Add(column);
                                    }
                                }
                                GloableVar.dataNum = rowCount;
                                GloableVar.whetherReadSuc = 1;
                                //填充行
                                for (int i = startRow; i <= rowCount; ++i)
                                {
                                    row = sheet.GetRow(i);
                                    if (row == null) continue;

                                    dataRow = dataTable.NewRow();
                                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                                    {
                                        cell = row.GetCell(j);
                                        if (cell == null)
                                        {
                                            dataRow[j] = "";
                                        }
                                        else
                                        {
                                            //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                            switch (cell.CellType)
                                            {
                                                case CellType.Blank:
                                                    dataRow[j] = "";
                                                    break;
                                                case CellType.Numeric:
                                                    short format = cell.CellStyle.DataFormat;
                                                    //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理
                                                    if (format == 14 || format == 31 || format == 57 || format == 58)
                                                        dataRow[j] = cell.DateCellValue;
                                                    else
                                                        dataRow[j] = cell.NumericCellValue;
                                                    break;
                                                case CellType.String:
                                                    dataRow[j] = cell.StringCellValue;
                                                    break;
                                            }
                                        }
                                    }
                                    dataTable.Rows.Add(dataRow);
                                    GloableVar.whetherReadSuc = 1;
                                }
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch (Exception)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                GloableVar.whetherReadSuc = 0;
                return null;
            }
        }

        public static void DataTableToExcel(string xlsxPath, DataTable dt)
        {
            bool result = false;
            IWorkbook workbook = null;
            FileStream fs = null;
            IRow row = null;
            ISheet sheet = null;
            ICell cell = null;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    workbook = new XSSFWorkbook();
                    sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表
                    int rowCount = dt.Rows.Count;//行数
                    int columnCount = dt.Columns.Count;//列数

                    //设置列头
                    row = sheet.CreateRow(0);//excel第一行设为列头
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }

                    //设置每行每列的单元格,
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行开始写入数据
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }

                    using (fs = File.OpenWrite(@xlsxPath))
                    {
                        workbook.Write(fs);//向打开的这个xls文件中写入数据
                        result = true;
                    }
                }
                //return result;
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                // return false;
            }
        }

    }

    public class GloableVar
    {
        public static int fringeNum = 0;
        public static int dataNum = 0;
        //public static int whetherRead = 0;
        public static int whetherReadSuc = 0;
        public static System.Windows.Threading.DispatcherTimer dTimer = new System.Windows.Threading.DispatcherTimer();
    }

    public class ExcelData
    {
        public static DataTable dataTable = new DataTable();
        public static string excelPath = "";
    }

}

