using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Color_Scales
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
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Create color scales for the data in specified range
                IConditionalFormats conditionalFormats = worksheet.Range["D7:D46"].ConditionalFormats;
                IConditionalFormat conditionalFormat = conditionalFormats.AddCondition();
                conditionalFormat.FormatType = ExcelCFType.ColorScale;
                IColorScale colorScale = conditionalFormat.ColorScale;

                //Sets 3 - color scale
                colorScale.SetConditionCount(3);
                colorScale.Criteria[0].FormatColorRGB = Color.FromArgb(230, 197, 218);
                colorScale.Criteria[0].Type = ConditionValueType.LowestValue;
                colorScale.Criteria[0].Value = "0";

                colorScale.Criteria[1].FormatColorRGB = Color.FromArgb(244, 210, 178);
                colorScale.Criteria[1].Type = ConditionValueType.Percentile;
                colorScale.Criteria[1].Value = "50";

                colorScale.Criteria[2].FormatColorRGB = Color.FromArgb(245, 247, 171);
                colorScale.Criteria[2].Type = ConditionValueType.HighestValue;
                colorScale.Criteria[2].Value = "0";

                conditionalFormat.FirstFormulaR1C1 = "=R[1]C[0]";
                conditionalFormat.SecondFormulaR1C1 = "=R[1]C[1]";

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}

