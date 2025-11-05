using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
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


namespace ConvertExcelToImage
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

        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open("Sample.xlsx");
                IWorksheet worksheet = workbook.Worksheets[0];

                //Convert the Excel to image
                System.Drawing.Image image = worksheet.ConvertToImage(1, 1, 20, 4);

                //Save the image as jpeg
                image.Save("Sample.Jpeg", ImageFormat.Jpeg);
            }

            //Launch the  Image        
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo = new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"../../Sample.Jpeg")) { UseShellExecute = true };
            process.Start();
            
        }
    }
}
