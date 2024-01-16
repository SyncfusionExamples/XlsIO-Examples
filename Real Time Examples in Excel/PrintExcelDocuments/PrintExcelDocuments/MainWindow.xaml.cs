using System.Drawing.Printing;
using System.Windows;
using Syncfusion.ExcelToPdfConverter;
using Syncfusion.XlsIO;

namespace PrintExcelDocuments
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

        private void SelectFile(object sender, RoutedEventArgs e)
        {
            //Initializes FileSavePicker.
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm,*.xltm,*.csv,*.tsv";
            openFileDialog.Title = "Select a Excel File";
            openFileDialog.ShowDialog();

            //Gets the path of specified file.
            filePath.Text = openFileDialog.FileName;            
        }

        private void PrintExcel(object sender, RoutedEventArgs e)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Loads or open an existing workbook through Open method of IWorkbooks
                IWorkbook workbook = application.Workbooks.Open(filePath.Text);

                //Initialize the printer settings
                PrinterSettings printerSettings = new PrinterSettings();

                //customizing the printer settings
                printerSettings.PrinterName = "HP LaserJet Pro MFP M127-M128 PCLmS";
                printerSettings.Copies = 2;
                printerSettings.FromPage = 2;
                printerSettings.ToPage = 3;
                printerSettings.DefaultPageSettings.Color = true;
                printerSettings.Duplex = Duplex.Vertical;
                printerSettings.Collate = true;

                ExcelToPdfConverter converter = new ExcelToPdfConverter(workbook);

                converter.Print();
            }
        }
    }
}