using System.IO;
using Syncfusion.XlsIO;

namespace Conditional_Format_with_R1C1
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

                //Using FormulaR1C1 property in Conditional Formatting 
                IConditionalFormats condition = worksheet.Range["E5:E18"].ConditionalFormats;
                IConditionalFormat condition1 = condition.AddCondition();
                condition1.FirstFormulaR1C1 = "=R[1]C[0]";
                condition1.SecondFormulaR1C1 = "=R[1]C[1]";

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ConditionalFormat.xlsx"));
                #endregion
            }
        }
    }
}




