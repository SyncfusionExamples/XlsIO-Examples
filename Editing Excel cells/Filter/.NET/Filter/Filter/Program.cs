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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
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
                workbook.SaveAs(Path.GetFullPath("Output/Filter.xlsx"));
                #endregion
            }
        }
    }
}




