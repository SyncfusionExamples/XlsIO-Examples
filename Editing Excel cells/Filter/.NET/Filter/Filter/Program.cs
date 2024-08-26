using System.IO;
using Syncfusion.XlsIO;

namespace Filter
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

                #region Filter
                //Creating an AutoFilter in the first worksheet. Specifying the AutoFilter range
                worksheet.AutoFilters.FilterRange = worksheet.Range["A1:A10"];

                //Column index to which AutoFilter must be applied
                IAutoFilter filter = worksheet.AutoFilters[0];

                //To apply Top10Number filter, IsTop and IsTop10 must be enabled
                filter.IsTop = true;
                filter.IsTop10 = true;

                //Setting Top10 filter with number of cell to be filtered from top
                filter.Top10Number = 5;
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Filter.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Filter.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
