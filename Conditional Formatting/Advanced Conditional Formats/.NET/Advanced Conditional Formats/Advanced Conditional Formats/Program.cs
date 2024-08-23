using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Advanced_Conditional_Formats
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

                //Create color scales for the data in specified range
                conditionalFormats = worksheet.Range["D7:D46"].ConditionalFormats;
                conditionalFormat = conditionalFormats.AddCondition();
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

                //Create icon sets for the data in specified range
                conditionalFormats = worksheet.Range["E7:E46"].ConditionalFormats;
                conditionalFormat = conditionalFormats.AddCondition();
                conditionalFormat.FormatType = ExcelCFType.IconSet;
                IIconSet iconSet = conditionalFormat.IconSet;

                //Apply three symbols icon and hide the data in the specified range
                iconSet.IconSet = ExcelIconSetType.ThreeSymbols;
                iconSet.IconCriteria[1].Type = ConditionValueType.Percent;
                iconSet.IconCriteria[1].Value = "50";
                iconSet.IconCriteria[2].Type = ConditionValueType.Percent;
                iconSet.IconCriteria[2].Value = "50";
                iconSet.ShowIconOnly = true;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("AdvancedCF.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}

