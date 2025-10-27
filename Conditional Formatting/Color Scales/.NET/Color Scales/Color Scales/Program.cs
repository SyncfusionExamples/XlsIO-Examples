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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);
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

                colorScale.Criteria[1].FormatColorRGB = Color.FromArgb(244, 210, 178);
                colorScale.Criteria[1].Type = ConditionValueType.Percentile;
                colorScale.Criteria[1].Value = "50";

                colorScale.Criteria[2].FormatColorRGB = Color.FromArgb(245, 247, 171);
                colorScale.Criteria[2].Type = ConditionValueType.HighestValue;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
                #endregion
            }
        }
    }
}





