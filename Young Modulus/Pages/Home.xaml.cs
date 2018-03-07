using FirstFloor.ModernUI.Windows.Controls;
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
using System.Windows.Threading;


namespace ModernUINavigationAppTest1.Pages
{
    /// <summary>
    /// Interaction logic for Home.xaml
    /// </summary>
    public partial class Home : UserControl
    {
        public Home()
        {
            InitializeComponent();
            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += timer_Tick;
            timer.Start();
            string exePath = Environment.CurrentDirectory;
            string folderPath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString() + "\\Young Modulus";
            try
            {
                CreateRunNumTxt(exePath);
                CreateResultTxt(exePath);
                CreateFriageTxt(exePath);
                CreateFolder(folderPath);
            }
            catch
            {
                MessageBox.Show("发生未知错误！", "Warnning!");
            }

            // = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString();
            //System.DateTime currentTime = new System.DateTime();
            //currentTime = System.DateTime.Now;
            //this.textBoxTime.Text = currentTime.ToString("yyyyMMdd_hhmmss");
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            this.textBlockShowTime.Text = DateTime.Now.ToString();
        }

        public void CreateFolder(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                DirectoryInfo dInfo = new DirectoryInfo(folderPath);
                dInfo.Create();
            }
        }

        public void CreateRunNumTxt(string exePath)
        {

            string txtPath = null;
            txtPath = exePath.Substring(0, exePath.Length) + "\\count.txt";
            if (File.Exists(txtPath))
            {
                // FileStream file = new FileStream(txtPath, FileMode.Open, FileAccess.ReadWrite);
                StreamReader sr = new StreamReader(txtPath, Encoding.Default);
                String line;
                line = sr.ReadLine();
                GlobalVar.runNum += Convert.ToInt32(line);
                sr.Close();
                TxtWrite(txtPath);
            }
            else
            {
                FileStream file = new FileStream(txtPath, FileMode.Create);
                GlobalVar.runNum = 1;
                StreamWriter sw = new StreamWriter(file);
                string str = Convert.ToString(GlobalVar.runNum);
                byte[] data = System.Text.Encoding.Default.GetBytes(str);
                file.Write(data, 0, data.Length);
                file.Flush();
                file.Close();
            }
        }

        public void TxtWrite(string txtPath)
        {
            FileStream file = new FileStream(txtPath, FileMode.Create);
            StreamWriter sw = new StreamWriter(file);
            string str = Convert.ToString(GlobalVar.runNum);
            byte[] data = System.Text.Encoding.Default.GetBytes(str);
            file.Write(data, 0, data.Length);
            file.Flush();
            file.Close();
        }

        public void CreateResultTxt(string exePath)
        {

            string txtPath = null;
            txtPath = exePath.Substring(0, exePath.Length) + "\\result.txt";
            if (!File.Exists(txtPath))
            {
                FileStream file = new FileStream(txtPath, FileMode.Create);
                StreamWriter sw = new StreamWriter(file);
                string str = "0";
                byte[] data = System.Text.Encoding.Default.GetBytes(str);
                file.Write(data, 0, data.Length);
                file.Flush();
                file.Close();
            }
        }

        public void CreateFriageTxt(string exePath)
        {

            string txtPath = null;
            txtPath = exePath.Substring(0, exePath.Length) + "\\friage.txt";
            if (!File.Exists(txtPath))
            {
                FileStream file = new FileStream(txtPath, FileMode.Create);
                StreamWriter sw = new StreamWriter(file);
                string str = "0";
                byte[] data = System.Text.Encoding.Default.GetBytes(str);
                file.Write(data, 0, data.Length);
                file.Flush();
                file.Close();
            }
        }

        //private void image_Loaded(object sender, RoutedEventArgs e)
        //{
        //    string path = Directory.GetCurrentDirectory();
        //    string path2 = "\\wedge interference.png";
        //    path = path.Substring(0, path.Length - 10) + path2;
        //    using (BinaryReader loader = new BinaryReader(File.Open(path, FileMode.Open)))
        //    {
        //        FileInfo fd = new FileInfo(path);
        //        int Length = (int)fd.Length;
        //        byte[] buf = new byte[Length];
        //        buf = loader.ReadBytes((int)fd.Length);
        //        loader.Dispose();
        //        loader.Close();


        //        //开始加载图像  
        //        BitmapImage bim = new BitmapImage();
        //        bim.BeginInit();
        //        bim.StreamSource = new MemoryStream(buf);
        //        bim.EndInit();
        //        image.Source = bim;
        //        GC.Collect(); //强制回收资源  
        //    }
        //}
    }
    public class GlobalVar
    {
        public static int runNum = 1;
    }
}
