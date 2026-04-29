using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Charts;
using System;

namespace ChartInvertIfNegative
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                #region Workbook Initialization
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];
                #endregion

                IChart chart = worksheet.Charts[0];

                //Used to invert series color if the value is negative
                (chart.Series[0] as ChartSerieImpl).InvertIfNegative = true;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
                #endregion
            }
        }
    }
}
