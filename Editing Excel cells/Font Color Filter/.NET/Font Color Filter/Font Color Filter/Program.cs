using System.IO;
using Syncfusion.XlsIO;

namespace Font_Color_Filter
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

                #region Font Color Filter
                //Creating an AutoFilter in the first worksheet. Specifying the AutoFilter range.
                worksheet.AutoFilters.FilterRange = worksheet.Range["A1:A11"];

                //Column index to which AutoFilter must be applied.
                IAutoFilter filter = worksheet.AutoFilters[0];

                //Applying color filter to filter based on Cell Color.
                filter.AddColorFilter(Syncfusion.Drawing.Color.Red, ExcelColorFilterType.FontColor);
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/FontColorFilter.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("FontColorFilter.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
