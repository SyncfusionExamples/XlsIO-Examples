using System;
using System.IO;
using Syncfusion.XlsIO;

namespace Time_Period
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Apply conditional format for specific time period
                IConditionalFormats conditionalFormats = worksheet.UsedRange.ConditionalFormats;
                IConditionalFormat conditionalFormat = conditionalFormats.AddCondition();

                //Set the format type to 'TimePeriod' to apply time-based conditional formatting
                conditionalFormat.FormatType = ExcelCFType.TimePeriod;
                conditionalFormat.TimePeriodType = CFTimePeriods.Today;

                //Set the background color of the matching cells 
                conditionalFormat.BackColor = ExcelKnownColors.Sky_blue;

                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }

    }
}
