using System.IO;
using Syncfusion.XlsIO;

namespace Place_In_Cell_Picture
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

                //Adding place in cell picture
                FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Image.png"), FileMode.Open, FileAccess.Read);
                IPictureShape picture = worksheet.Pictures.AddPicture(1, 1, imageStream);

                picture.PlaceInCell = true;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/PlaceInCellPicture.xlsx"));
                #endregion

                //Dispose streams
                imageStream.Dispose();
            }
        }
    }
}





