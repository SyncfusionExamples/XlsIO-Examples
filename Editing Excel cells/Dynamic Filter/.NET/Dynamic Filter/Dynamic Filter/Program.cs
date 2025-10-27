using System.IO;
using Syncfusion.XlsIO;

namespace Dynamic_Filter
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

                #region Dynamic Filter
                //Creating an AutoFilter in the first worksheet. Specifying the AutoFilter range.
                worksheet.AutoFilters.FilterRange = worksheet.Range["A1:A13"];

                //Column index to which AutoFilter must be applied.
                IAutoFilter filter = worksheet.AutoFilters[0];

                //Applying dynamic filter to filter the date based on DynamicFilterType.
                filter.AddDynamicFilter(DynamicFilterType.NextQuarter);
				#endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/DynamicFilter.xlsx"));
                #endregion
            }
        }
    }
}




