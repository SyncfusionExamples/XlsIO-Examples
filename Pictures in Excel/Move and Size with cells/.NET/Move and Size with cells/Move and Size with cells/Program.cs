using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;

namespace Move_and_Size_with_cells
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding a picture
                FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Image.png"), FileMode.Open, FileAccess.Read);
                IPictureShape shape = worksheet.Pictures.AddPicture(1, 1, 5, 3, imageStream);
                shape = worksheet.Pictures.AddPicture(1, 5, 5, 7, imageStream);
                
                //Set move picture with cell
                shape.IsMoveWithCell = true;

                //Set size picture with cell
                shape.IsSizeWithCell = true;

                //Hide the column
                worksheet.HideColumn(5);

                //Saving the workbook as stream
                FileStream OutputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(OutputStream);

                //Dispose streams
                imageStream.Dispose();
                OutputStream.Dispose();
            }
        }
    }
}




