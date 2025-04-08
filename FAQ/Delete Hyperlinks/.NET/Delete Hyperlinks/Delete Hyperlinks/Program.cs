using System;
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;
using Syncfusion.XlsIO.Implementation.Collections;

namespace Create_Table
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];
                // Remove first hyperlink without affecting cell styles
                HyperLinksCollection hyperlink = worksheet.HyperLinks as HyperLinksCollection;
                hyperlink.Remove(hyperlink[0] as HyperLinkImpl);
                FileStream outputStream = new FileStream("Output/Output.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                workbook.Close();
                excelEngine.Dispose();
            }
        }

    }
}

