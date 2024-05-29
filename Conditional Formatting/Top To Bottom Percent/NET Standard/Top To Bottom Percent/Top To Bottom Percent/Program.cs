using System.IO;
using Syncfusion.XlsIO;

namespace Top_To_Bottom_Percent
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
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Applying conditional formatting to "N6:N35".
                IConditionalFormats conditionalFormats1 = worksheet.Range["N6:N35"].ConditionalFormats;
                IConditionalFormat conditionalFormat1 = conditionalFormats1.AddCondition();

                //Applying top or bottom rule in the conditional formatting.
                conditionalFormat1.FormatType = ExcelCFType.TopBottom;
                ITopBottom topBottom1 = conditionalFormat1.TopBottom;

                //Set type as Bottom for TopBottom rule.
                topBottom1.Type = ExcelCFTopBottomType.Bottom;

                //Set true to Percent property for TopBottom rule.
                topBottom1.Percent = true;

                //Set rank value for the TopBottom rule.
                topBottom1.Rank = 50;

                //Set solid color conditional formatting for TopBottom rule.
                conditionalFormat1.FillPattern = ExcelPattern.Solid;
                conditionalFormat1.BackColorRGB = Syncfusion.Drawing.Color.FromArgb(51, 153, 102);

                //Applying conditional formatting to "M6:M35".
                IConditionalFormats conditionalFormats2 = worksheet.Range["M6:M35"].ConditionalFormats;
                IConditionalFormat conditionalFormat2 = conditionalFormats2.AddCondition();

                //Applying top or bottom rule in the conditional formatting.
                conditionalFormat2.FormatType = ExcelCFType.TopBottom;
                ITopBottom topBottom2 = conditionalFormat2.TopBottom;

                //Set type as Top for TopBottom rule.
                topBottom2.Type = ExcelCFTopBottomType.Bottom;

                //Set true to Percent property for TopBottom rule.
                topBottom2.Percent = true;

                //Set rank value for the TopBottom rule.
                topBottom2.Rank = 20;

                //Set gradient color conditional formatting for TopBottom rule.
                conditionalFormat2.FillPattern = ExcelPattern.Gradient;
                conditionalFormat2.BackColorRGB = Syncfusion.Drawing.Color.FromArgb(130, 60, 12);
                conditionalFormat2.ColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 255, 0);
                conditionalFormat2.GradientStyle = ExcelGradientStyle.Horizontal;
                conditionalFormat2.GradientVariant = ExcelGradientVariants.ShadingVariants_1;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("TopToBottomRank.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("TopToBottomRank.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
