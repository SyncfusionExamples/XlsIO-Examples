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
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
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
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/AdvancedFilter.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}




