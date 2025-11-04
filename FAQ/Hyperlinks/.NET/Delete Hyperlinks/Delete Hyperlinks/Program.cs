using System;
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;
using Syncfusion.XlsIO.Implementation.Collections;

namespace Delete_Hyperlinks
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                // Remove first hyperlink without affecting cell styles
                HyperLinksCollection hyperlink = worksheet.HyperLinks as HyperLinksCollection;
                hyperlink.Remove(hyperlink[0] as HyperLinkImpl);

                //Saving the workbook 
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
                workbook.Close();
                excelEngine.Dispose();
            }
        }
    }
}