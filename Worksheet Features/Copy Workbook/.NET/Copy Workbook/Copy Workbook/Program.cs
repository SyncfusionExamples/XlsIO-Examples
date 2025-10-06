using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Collections;
using Syncfusion.XlsIO.Implementation;
using Syncfusion.XlsIO.Interfaces;

namespace Copy_Workbook
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook sourceWorkbook = application.Workbooks.Open(Path.GetFullPath(@"Data/SourceWorkbookTemplate.xlsx"));
                IWorkbook destinationWorkbook = application.Workbooks.Open(Path.GetFullPath(@"Data/DestinationWorkbookTemplate.xlsx"));

                //Clone the workbook
                destinationWorkbook = sourceWorkbook.Clone();
               
                //Saving the workbook
                destinationWorkbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }   
        }
    }
}




