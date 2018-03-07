using FirstFloor.ModernUI.Windows.Controls;
using FirstFloor.ModernUI;
using FirstFloor.ModernUI.App;
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


/* ChangeLog
 * 2017.2.25 √实现后台计算代码，数据从1.txt里面读取（每行一个数据）
 * 2017.3.3   √增加读取数据button，添加说明信息
 * 2017.3.4   √增加初始化按钮以及关闭按钮
 * 2017.3.5   √增加Slider，使得条纹数目可以拖动而改变，从而实现两种方式
 * 2017.3.6   √尝试使用List存储数据，并且实现能够计算；
 * 2017.3.7   √实现List与DataGrid数据绑定
 * 2017.3.8   √增加了全局使用的变量类GlobalVar，使得程序更加方便
 * 2017.3.9   √使用Microsoft.ACE.OLEDB.12.0和using System.Data.OleDb;来读取数据，但是需要其他组件
 * 2017.3.11 √增加了NPOI引用
 * 2017.3.13 √成功实现从Excel读取数据并计算，下一步是实现DataGrid与DataTable双向绑定
                      增加ExcelData类，用DataTable来存储从Excel读入的数据，并且再从DataTable中读取数据进行计算
 * 2017.3.14 √实现判断计算前是否读取数据，并且判断数据是否成功读取
                    √实现DataGrid与DataTable数据双向绑定     
 * 2017.3.18 √增加在DataGrid 中修改数据并更新数据
 * 2017.3.19 √增加第一次进入程序就能在DataGrid中输入数据(FirstRun函数)，并且删除是否读取数据判断
                    √增加如果文件地址中含有"",那么在程序中自动去掉
 * 2017.3.20 √修改了判断能否执行计算的条件，增加了部分说明界面
 * 2017.3.21 √增加关于中的说明和更新日志
 * 2017.3.22 √界面内容优化
 * 2017.3.23 √增加选择数据按钮，可以自己选择Excel文件位置
                      更换软件图标           
 * 2017.3.26   删除不必要的引用和代码
 * 2017.3.27   界面优化
 * 2017.3.28  √增加统计信息界面，包含启动次数，程序版本，上次计算结果；主页优化
 * 2017.3.29  √增加计算--上次计算数据，并且可以实时更新与导出；优化统计界面信息
 *                     导入数据与显示上次数据用到了NPOI导出Excel以及读取写入txt
 *                     导出数据时可以自定义路径与文件名
 *                     显示程序路径，并且可以在资源管理器中打开
 * 2017.3.30   √自动保存每一次计算的数据至用户文档Young Modulus文件夹，命名格式：时间_条纹数目_计算结果.xlsx
 *                      统计信息中增加打开Young Modulus文件夹，显示这个文件夹的文件数目以及大小，增加清理缓存按钮
 * 2017.4.1     √优化界面UI；增加定时器Timer(); 使得显示  读取成功！ 导出成功！ 5秒后自动消失    
 * 2017.4.5     √将部分TextBox替换为TextBlock，避免发生不必要的修改
 * 2017.5.4     √优化界面UI，修改屏幕分辨率的计算逻辑，增加显示系统DPI缩放
 * 2017.5.6     √修复了在文本框输入条纹数目不起作用的bug，加大滑动条区间
*/

namespace Young_Modulus
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : ModernWindow
    {
        public MainWindow()
        {
            
            InitializeComponent();
        }
    }
}

