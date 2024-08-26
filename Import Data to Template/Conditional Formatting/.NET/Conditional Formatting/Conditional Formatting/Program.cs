using Syncfusion.Drawing;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.IO;

namespace Conditional_Formatting
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

                //Create Template Marker Processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                IConditionalFormats conditionalFormats = marker.CreateConditionalFormats(worksheet["C5"]);

                #region Data Bar

                //Apply markers using Formula
                IConditionalFormat condition = conditionalFormats.AddCondition();

                //Set Data bar and icon set for the same cell
                //Set the format type
                condition.FormatType = ExcelCFType.DataBar;

                IDataBar dataBar = condition.DataBar;

                //Set the constraint
                dataBar.MinPoint.Type = ConditionValueType.LowestValue;
                dataBar.MinPoint.Value = "0";
                dataBar.MaxPoint.Type = ConditionValueType.HighestValue;
                dataBar.MaxPoint.Value = "0";

                //Set color for Bar
                dataBar.BarColor = Color.FromArgb(156, 208, 243);

                //Hide the value in data bar
                dataBar.ShowValue = false;

                #endregion

                #region IconSet

                //Declaring Icon set for the condition
                condition = conditionalFormats.AddCondition();
                condition.FormatType = ExcelCFType.IconSet;

                IIconSet iconSet = condition.IconSet;
                iconSet.IconSet = ExcelIconSetType.FourRating;
                iconSet.IconCriteria[0].Type = ConditionValueType.LowestValue;
                iconSet.IconCriteria[0].Value = "0";
                iconSet.IconCriteria[1].Type = ConditionValueType.HighestValue;
                iconSet.IconCriteria[1].Value = "0";
                iconSet.ShowIconOnly = true;

                #endregion

                conditionalFormats = marker.CreateConditionalFormats(worksheet["D5"]);

                #region Color Scale

                //Applying color for Conditional formatting
                condition = conditionalFormats.AddCondition();
                condition.FormatType = ExcelCFType.ColorScale;

                IColorScale colorScale = condition.ColorScale;

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

                #endregion

                conditionalFormats = marker.CreateConditionalFormats(worksheet["E5"]);

                #region IconSet

                //Apply Icons for conditional formatting
                condition = conditionalFormats.AddCondition();
                condition.FormatType = ExcelCFType.IconSet;

                iconSet = condition.IconSet;
                iconSet.IconSet = ExcelIconSetType.ThreeSymbols;

                iconSet.IconCriteria[0].Type = ConditionValueType.LowestValue;
                iconSet.IconCriteria[0].Value = "0";

                iconSet.IconCriteria[1].Type = ConditionValueType.HighestValue;
                iconSet.IconCriteria[1].Value = "0";

                iconSet.ShowIconOnly = false;

                #endregion

                //Add collection to the marker variables where the name should match with input template
                marker.AddVariable("Reports", GetSalesReports());

                //Process the markers in the template
                marker.ApplyMarkers();


                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ConditionalFormatting.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
        public static IList<Sales> GetSalesReports()
        {
            IList<Sales> sales = new List<Sales>();

            sales.Add(new Sales("Andy Bernard", 45000, 58000, 29));

            sales.Add(new Sales("Jim Halpert", 34000, 65000, 91));

            sales.Add(new Sales("Karen Fillippelli", 75000, 64000, -15));

            sales.Add(new Sales("Phyllis Lapin", 56500, 33600, -40));

            sales.Add(new Sales("Stanley Hudson", 46500, 52000, 12));

            return sales;
        }
        public class Sales
        {
            public string SalesPerson { get; set; }
            public int SalesJanJun { get; set; }
            public int SalesJulDec { get; set; }
            public int Change { get; set; }

            public Sales(string name, int salesJanJun, int salesJulDec, int change)
            {
                SalesPerson = name;
                SalesJanJun = salesJanJun;
                SalesJulDec = salesJulDec;
                Change = change;
            }
        }
    }
}





