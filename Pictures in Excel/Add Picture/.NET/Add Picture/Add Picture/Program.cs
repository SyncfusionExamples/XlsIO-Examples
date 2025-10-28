using System.IO;
using Syncfusion.XlsIO;

namespace Add_Picture
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

                //Adding a picture
                FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Image.png"), FileMode.Open, FileAccess.Read);
                IPictureShape shape = worksheet.Pictures.AddPicture(1, 1, imageStream);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/AddPicture.xlsx"));
                #endregion

                //Dispose streams
                imageStream.Dispose();
            }
        }
    }
}





