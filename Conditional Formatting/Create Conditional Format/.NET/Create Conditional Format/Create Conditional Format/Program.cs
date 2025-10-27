using System.IO;
using Syncfusion.XlsIO;

namespace Create_Conditional_Format
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Applying conditional formatting to "F2"
                IConditionalFormats condition = worksheet.Range["F2"].ConditionalFormats;
                IConditionalFormat condition1 = condition.AddCondition();

                //Represents conditional format rule that the value in target range should be between 10 and 20
                condition1.FormatType = ExcelCFType.CellValue;
                condition1.Operator = ExcelComparisonOperator.Between;
                condition1.FirstFormula = "10";
                condition1.SecondFormula = "20";
                worksheet.Range["A2"].Text = "Enter a number between 10 and 20";
                worksheet.Range["F2"].BorderAround(ExcelLineStyle.Thin);

                //Setting back color and font style to be applied for target range
                condition1.BackColor = ExcelKnownColors.Light_orange;
                condition1.IsBold = true;
                condition1.IsItalic = true;

                //Applying conditional formatting to "F4"
                condition = worksheet.Range["F4"].ConditionalFormats;
                IConditionalFormat condition2 = condition.AddCondition();

                //Represents conditional format rule that the cell value should be 1000
                condition2.FormatType = ExcelCFType.CellValue;
                condition2.Operator = ExcelComparisonOperator.Equal;
                condition2.FirstFormula = "1000";
                worksheet.Range["A4"].Text = "Enter the Number as 1000";
                worksheet.Range["F4"].BorderAround(ExcelLineStyle.Thin);

                //Setting fill pattern and back color to target range
                condition2.FillPattern = ExcelPattern.LightUpwardDiagonal;
                condition2.BackColor = ExcelKnownColors.Yellow;

                //Applying conditional formatting to "F6"
                condition = worksheet.Range["F6"].ConditionalFormats;
                IConditionalFormat condition3 = condition.AddCondition();

                //Setting conditional format rule that the cell value for target range should be less than or equal to 1000
                condition3.FormatType = ExcelCFType.CellValue;
                condition3.Operator = ExcelComparisonOperator.LessOrEqual;
                condition3.FirstFormula = "1000";
                worksheet.Range["A6"].Text = "Enter a Number which is less than or equal to 1000";
                worksheet.Range["F6"].BorderAround(ExcelLineStyle.Thin);

                //Setting back color to target range
                condition3.BackColor = ExcelKnownColors.Light_green;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ConditionalFormat.xlsx"));
                #endregion
            }
        }
    }
}