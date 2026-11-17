using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace NamedRange_to_Image
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Sample.xlsb"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                IName name = workbook.Names["Range1"];          

                IRange range = name.RefersToRange;

                //Initialize XlsIORenderer
                application.XlsIORenderer = new XlsIORenderer();

                #region Save
                //Saving the workbook
                using FileStream outputStream = new FileStream("../../../Output/Output.png", FileMode.Create, FileAccess.Write);

                // Convert the named range to an image.
                range.Worksheet.ConvertToImage(range, outputStream);

                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}