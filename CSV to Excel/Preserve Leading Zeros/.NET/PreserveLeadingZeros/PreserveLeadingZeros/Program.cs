using Syncfusion.XlsIO;

namespace PreserveLeadingZeros
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                application.PreserveCSVDataTypes = true;

                //Enable KeepLeadingZeros property 
                application.KeepLeadingZeros = true;

                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.csv"), ",");
                
                //Save the workbook 
                workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
            }
        }
    }
}