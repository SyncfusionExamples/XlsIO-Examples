using System.IO;
using Syncfusion.XlsIO;

namespace Icon_Filter
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Icon Filter
                //Creating an AutoFilter in the first worksheet. Specifying the AutoFilter range.
                worksheet.AutoFilters.FilterRange = worksheet.Range["A1:A8"];

                //Column index to which AutoFilter must be applied.
                IAutoFilter filter = worksheet.AutoFilters[0];

                //Applying Icon filter to filter based on applied icon set.
                filter.AddIconFilter(ExcelIconSetType.ThreeFlags, 2);
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("IconFilter.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("IconFilter.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
