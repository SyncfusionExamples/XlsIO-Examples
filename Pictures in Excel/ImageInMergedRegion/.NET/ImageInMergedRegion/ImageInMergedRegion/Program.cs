using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Shapes;
using System.IO;

namespace ImageInMergedRegion
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Get the merged cells
                IRange[] range = new IRange[3];
                range[0] = worksheet.MergedCells[0];
                range[1] = worksheet.MergedCells[1];
                range[2] = worksheet.MergedCells[2];

                //Get the images
                string[] image = new string[3];
                image[0] = Path.GetFullPath(@"Data/Picture1.png");
                image[1] = Path.GetFullPath(@"Data/Picture2.png");
                image[2] = Path.GetFullPath(@"Data/Picture3.png");

                //Insert images
                int i = 0;
                foreach (IRange cell in range)
                {
                    FileStream imageStream = new FileStream(image[i], FileMode.Open, FileAccess.Read);
                    IPictureShape shape = worksheet.Pictures.AddPicture(cell.Row, cell.Column, imageStream);
                    (shape as ShapeImpl).BottomRow = cell.MergeArea.LastRow;
                    (shape as ShapeImpl).RightColumn = cell.MergeArea.LastColumn;
                    i++;
                    imageStream.Dispose();
                }

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("ImageInMergedRegion.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

            }
        }
    }
}

