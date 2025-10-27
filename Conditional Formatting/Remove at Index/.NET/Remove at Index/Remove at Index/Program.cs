using System.IO;
using Syncfusion.XlsIO;

namespace Remove_at_Index
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Removing first conditional Format at the specified Range
                worksheet.UsedRange.ConditionalFormats.RemoveAt(0);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/RemoveConditionalFormat.xlsx"));
                #endregion
            }
        }
    }
}





