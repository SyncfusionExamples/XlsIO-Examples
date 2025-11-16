using System;
using System.IO;
using Syncfusion.XlsIO;

namespace Traffic_Light
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Add data and formatting to the worksheet
                worksheet.Range["A1"].Text = "Traffic Lights";

                //Add percentage values to cells A2 to A7 and format them as percentages
                worksheet.Range["A2"].Number = 0.95;
                worksheet.Range["A2"].NumberFormat = "0%";
                worksheet.Range["A3"].Number = 0.5;
                worksheet.Range["A3"].NumberFormat = "0%";
                worksheet.Range["A4"].Number = 0.1;
                worksheet.Range["A4"].NumberFormat = "0%";
                worksheet.Range["A5"].Number = 0.9;
                worksheet.Range["A5"].NumberFormat = "0%";
                worksheet.Range["A6"].Number = 0.7;
                worksheet.Range["A6"].NumberFormat = "0%";
                worksheet.Range["A7"].Number = 0.6;
                worksheet.Range["A7"].NumberFormat = "0%";

                //Adjust row height and column width of the used range
                worksheet.UsedRange.RowHeight = 20;
                worksheet.UsedRange.ColumnWidth = 25;

                //Apply the first conditional format
                IConditionalFormats condition = worksheet.UsedRange.ConditionalFormats;
                IConditionalFormat condition1 = condition.AddCondition();

                condition1.FormatType = ExcelCFType.CellValue;
                condition1.FirstFormula = "300";
                condition1.Operator = ExcelComparisonOperator.Less;
                condition1.FontColor = ExcelKnownColors.Black;
                condition1.BackColor = ExcelKnownColors.Sky_blue;

                //Apply the second conditional format
                IConditionalFormats condition2 = worksheet.UsedRange.ConditionalFormats;
                IConditionalFormat condition3 = condition2.AddCondition();
                condition3.FormatType = ExcelCFType.IconSet;
                IIconSet iconSet = condition3.IconSet;
                iconSet.IconSet = ExcelIconSetType.ThreeTrafficLights1;

                //Saving the workbook
                workbook.SaveAs("Output.xlsx");
            }
        }
    }
}


