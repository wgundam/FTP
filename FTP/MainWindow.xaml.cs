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

namespace FTP
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        int count = 1;
        public MainWindow()
        {
            InitializeComponent();

        }

        private void Clear(object sender, RoutedEventArgs e)
        {
            count = 1;
            Init();
            if (File.Exists(@"C:\YYK\compare.xls"))
                File.Delete(@"C:\YYK\compare.xls");
            OUTPUT.Text = "缓存删除成功！";
        }

        private void Generate(object sender, RoutedEventArgs e)
        {
            try
            {
                Init();
                Process();
                OUTPUT.Text = "第" + count + "次成功生成！";
                count++;
            }
            catch (Exception ex)
            { 
                MessageBox.Show(ex.ToString()); 
            }
        }

        //处理
        public void Process()
        {
            //读取FTP信息列表
            string GetFileList = FTPMethod.GetFileList();
            //Console.WriteLine(GetFileList);
            //根据列表下载所需文件
            FTPMethod.FileFilter(GetFileList);
            //生产EXCEL
            FileMethod.ExcelToday();
            //上传至FTP
            FTPMethod.UploadFolder();
        }

        //启动软件，清空历史
        public static void Init()
        {
            if (!Directory.Exists(@"C:\YYK"))
                Directory.CreateDirectory(@"C:\YYK");
            if (Directory.Exists(@"C:\YYK\YYK_OUTPUT"))
                Directory.Delete(@"C:\YYK\YYK_OUTPUT",true);
            if (!Directory.Exists(@"C:\YYK\YYK_OUTPUT"))
                Directory.CreateDirectory(@"C:\YYK\YYK_OUTPUT\");
            if (Directory.Exists(@"C:\YYK\YYK_TEMP"))
                Directory.Delete(@"C:\YYK\YYK_TEMP",true);
            if (!Directory.Exists(@"C:\YYK\YYK_TEMP"))
                Directory.CreateDirectory(@"C:\YYK\YYK_TEMP\");
            
        }

        private void window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Environment.Exit(0);
        }


    }
}
