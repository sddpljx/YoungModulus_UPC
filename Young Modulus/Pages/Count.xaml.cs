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
using System.Text;
using System.Deployment.Application;
using System.Data;
using System.Reflection;
using System.Drawing;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.Win32;

namespace Young_Modulus.Pages
{
    /// <summary>
    /// Interaction logic for Count.xaml
    /// </summary>
    public partial class Count : UserControl
    {
        public Count()
        {
            InitializeComponent();
            string str;
            str = GetVersion() + ".rs2_release";
            this.textBlockShowVersion.Text = str;
            ShowRunNum();
            ShowSoftwareAddr();
            GetFolderInfo();
            GetScreenSquare();
            GetDotNetVersion();
            //    this.textBoxShowOSVer.Text = Environment.OSVersion.ToString();
        }

        public void ShowRunNum()
        {
            string exePath = Environment.CurrentDirectory;
            string txtPath = null;
            txtPath = exePath.Substring(0, exePath.Length) + "\\count.txt";
            StreamReader sr = new StreamReader(txtPath, Encoding.Default);
            String line;
            line = sr.ReadLine();
            this.textBlockShowRunNum.Text = line;
            sr.Close();
        }

        public string GetVersion()
        {
            //return App.ResourceAssembly.GetName(false).Version.ToString();
            try
            {
                return ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            catch
            {
                return "Debug Mode";
            }
        }

        public void ShowSoftwareAddr()
        {
            string exePath = Environment.CurrentDirectory;
            this.textBlockShowSoftAddr.Text = exePath;
        }

        private void buttonOpenSoftAddr_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe ", Environment.CurrentDirectory);
        }

        private void buttonOpenDataAddr_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString() + "\\Young Modulus";
            System.Diagnostics.Process.Start("explorer.exe ", folderPath);
        }

        public void GetFolderInfo()
        {
            try
            {
                string folderPath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString() + "\\Young Modulus";
                DirectoryInfo dI = new DirectoryInfo(folderPath);
                long len = 0;
                int cnt = 0;
                double lenKB = 0;
                foreach (FileInfo fi in dI.GetFiles())
                {
                    len += fi.Length;
                    cnt++;
                }
                if (len <= 1024 * 1024 & len >= 0)
                {
                    lenKB = 1.0 * len / 1024;
                    this.textBlockFolderLen.Text = "    目录占用空间：" + Convert.ToString(Math.Round(lenKB, 2)) + "KB";
                }
                else if (len >= 1024 * 1024)
                {
                    lenKB = 1.0 * len / (1024 * 1024);
                    this.textBlockFolderLen.Text = "    目录占用空间：" + Convert.ToString(Math.Round(lenKB, 2)) + "MB";
                }
                else
                {
                    lenKB = 1.0 * len / (1024 * 1024 * 1024);
                    this.textBlockFolderLen.Text = "    目录占用空间：" + Convert.ToString(Math.Round(lenKB, 2)) + "GB";
                }
                this.textBlockFolderCnt.Text = "    目录文件数量：" + cnt.ToString();

                //if (len >= 50*1024*1024)
                //{
                //    CleanFolder();
                //}
            }
            catch
            {

            }
        }

        private void buttonCleanFolder_Click(object sender, RoutedEventArgs e)
        {
            CleanFolder();
        }

        public void CleanFolder()
        {
            string folderPath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString() + "\\Young Modulus";
            DirectoryInfo di = new DirectoryInfo(folderPath);
            di.Delete(true);
            di.Create();
            GetFolderInfo();
        }

        public void GetScreenSquare()
        {
            int width = Convert.ToInt32(SystemParameters.PrimaryScreenWidth);
            int height = Convert.ToInt32(SystemParameters.PrimaryScreenHeight);
            double systemDPI = 1;
            double width1 = 0, height1 = 0;
            Graphics graphics = Graphics.FromHwnd(IntPtr.Zero);
            double dPIX = graphics.DpiX;
            double dPIY = graphics.DpiY;
            if (dPIX <= 100)
            {
                systemDPI = 1;
            }
            else if (dPIX <= 125)
            {
                systemDPI = 1.25;
            }
            else if (dPIX <= 150)
            {
                systemDPI = 1.5;
            }
            else if (dPIX <= 175)
            {
                systemDPI = 1.75;
            }
            else if (dPIX <= 200)
            {
                systemDPI = 2;
            }
            width1 = width * systemDPI;
            height1 = height * systemDPI;
            width = Convert.ToInt32(Math.Ceiling(width1));
            height = Convert.ToInt32(Math.Ceiling(height1));
            string squ = width.ToString() + "x" + height.ToString();
            this.textBlockShowScreenSqu.Text = "屏幕分辨率：   " + squ + "     DPI: " + ((int)(systemDPI * 100)).ToString() + "%";
        }

        public string GetDotNetVersion()
        {
            try
            {
                RegistryKey reg = null;
                reg = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full", false);
                int releaseKey = Convert.ToInt32(reg.GetValue("Release"));
                string ver = null;
                if (releaseKey != null)
                {
                    if (releaseKey >= 460798)
                    {
                        ver = "4.7";
                    }
                    else if (releaseKey >= 394802)
                    {
                        ver = "4.6.2";
                    }
                    else if (releaseKey >= 394254)
                    {
                        ver = "4.6.1";
                    }
                    else if (releaseKey >= 393295)
                    {
                        ver = "4.6";
                    }
                    else if (releaseKey >= 379893)
                    {
                        ver = "4.5.2";
                    }
                    else if (releaseKey >= 378758)
                    {
                        ver = "4.5.1";
                    }
                    else if (releaseKey >= 378389)
                    {
                        ver = "4.5";
                    }
                    this.textBlockDotNetVer.Text = "Microsoft .NET FrameWork " + ver;
                }
                reg.Close();
                return ver;
            }
            catch
            {
                return null;
            }
        }

    }

}
