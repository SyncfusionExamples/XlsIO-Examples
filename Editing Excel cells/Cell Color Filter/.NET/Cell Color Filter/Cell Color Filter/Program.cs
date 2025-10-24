using System.IO;
using Syncfusion.XlsIO;

namespace Cell_Color_Filter
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

                #region Cell Color Filter
                //Creating an AutoFilter in the first worksheet. Specifying the AutoFilter range.
                worksheet.AutoFilters.FilterRange = worksheet.Range["A1:A11"];

                //Column index to which AutoFilter must be applied.
                IAutoFilter filter = worksheet.AutoFilters[0];

                //Applying color filter to filter based on Cell Color.
                filter.AddColorFilter(Syncfusion.Drawing.Color.Red, ExcelColorFilterType.CellColor);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/CellColorFilter.xlsx"));
                #endregion
            }
        }
    }
}




