using Syncfusion.XlsIO;

namespace ColumnNumberFormat
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));
                IWorksheet sheet = workbook.Worksheets[0];

                //Case 1: Apply direct number format (zero-based index)
                sheet.Columns[0].NumberFormat = "yyyy-mm-dd"; //Column A
                sheet.Columns[3].NumberFormat = "$#,##0.00"; //Column D
                sheet.Columns[4].NumberFormat = "0.00%"; //Column E

                //Case 2: Apply style-based format (one-based index)
                IStyle style = workbook.Styles.Add("DecimalStyle");
                style.NumberFormat = "0.00";
                sheet.SetDefaultColumnStyle(3, style); //Column C 

                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
            }
        }
    }
}