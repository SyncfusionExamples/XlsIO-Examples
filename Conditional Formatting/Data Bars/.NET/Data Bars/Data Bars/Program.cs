using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Data_Bars
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Create data bars for the data in specified range
                IConditionalFormats conditionalFormats = worksheet.Range["C7:C46"].ConditionalFormats;
                IConditionalFormat conditionalFormat = conditionalFormats.AddCondition();
                conditionalFormat.FormatType = ExcelCFType.DataBar;
                IDataBar dataBar = conditionalFormat.DataBar;

                //Set the constraints
                dataBar.MinPoint.Type = ConditionValueType.LowestValue;
                dataBar.MaxPoint.Type = ConditionValueType.HighestValue;

                //Set color for Bar
                dataBar.BarColor = Color.FromArgb(156, 208, 243);

                //Hide the values in data bar
                dataBar.ShowValue = false;
                dataBar.BarColor = Color.Aqua;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Output.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
