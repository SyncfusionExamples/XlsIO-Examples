using System.IO;
using Syncfusion.XlsIO;

namespace External_Image
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Add image from the specified url at the specified location in the worksheet
                worksheet.Pictures.AddPictureAsLink(1, 1, 5, 7, "https://cdn.syncfusion.com/content/images/company-logos/Syncfusion_Logo_Image.png");

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ExternalImage.xlsx"));
                #endregion
            }
        }
    }
}




