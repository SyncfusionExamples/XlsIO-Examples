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
using Syncfusion.XlsIO;

namespace Edit_Excel
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Set the default application version
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load the existing Excel workbook into IWorkbook
                IWorkbook workbook = application.Workbooks.Open("../../InputTemplate.xlsx");

                //Get the first worksheet in the workbook into IWorksheet
                IWorksheet worksheet = workbook.Worksheets[0];

                //Assign some text in a cell
                worksheet.Range["A3"].Text = "Hello World";

                //Save the Excel document
                workbook.SaveAs("EditExcel.xlsx");
                System.Diagnostics.Process.Start("EditExcel.xlsx");
            }
        }
    }
}
