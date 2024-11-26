using System.IO;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;

namespace Header_and_Footer
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding values in worksheet
                worksheet.Range["A1:A600"].Text = "HelloWorld";

                //Adding text with red color formatting to the left header
                worksheet.PageSetup.LeftHeader = "&KFF0000 Left Header";

                //Adding text (sheet name) with red color formatting to the center header
                worksheet.PageSetup.CenterHeader = "&KFF0000&A";

                //Adding text (date and time) with red color formatting to the right header
                worksheet.PageSetup.RightHeader = "&KFF0000&D &T";

                //Adding bold text, font size 18, and blue color formatting to the left footer
                worksheet.PageSetup.LeftFooter = "&B &18 &K0000FF Left Footer";

                //Adding an image placeholder to the center footer
                FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/image.jpg"), FileMode.Open);
                worksheet.PageSetup.CenterFooter = "&G";
                worksheet.PageSetup.CenterFooterImage = Image.FromStream(imageStream);

                //Adding the file name to the right footer with blue color formatting
                worksheet.PageSetup.RightFooter = "&K0000FF&F";

                //Saving the workbook as stream
                FileStream stream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create);
                workbook.SaveAs(stream);
                stream.Dispose();
            }
        }
    }
}





