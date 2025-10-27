using System.IO;
using Syncfusion.XlsIO;

namespace Remove_all_Conditional_Formats
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

                //Removing Conditional Formatting Settings From Entire Sheet
                worksheet.UsedRange.Clear(ExcelClearOptions.ClearConditionalFormats);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/RemoveAll.xlsx"));
                #endregion
            }
        }
    }
}





