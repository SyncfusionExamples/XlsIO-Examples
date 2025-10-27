using System.IO;
using Syncfusion.XlsIO;

namespace Advanced_Filter
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

                #region Advanced Filter
                IRange filterRange = worksheet.Range["A8:G51"];
                IRange criteriaRange = worksheet.Range["A2:B5"];
                IRange copyToRange = worksheet.Range["I8"];

                //Apply the Advanced Filter with enable of unique value and copy to another place.
                worksheet.AdvancedFilter(ExcelFilterAction.FilterCopy, filterRange, criteriaRange, copyToRange, true);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/AdvancedFilter.xlsx"));
                #endregion
            }
        }
    }
}




