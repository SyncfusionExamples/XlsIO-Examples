using Syncfusion.XlsIO;
using System.IO;

class Program
{
    public static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"../../../Data/Input.xlsx"));
            IWorksheet worksheet = workbook.Worksheets[0];

            //Creating an AutoFilter 
            worksheet.AutoFilters.FilterRange = worksheet.Range["A1:A11"];
            IAutoFilter filter = worksheet.AutoFilters[0];

            //Specifying first condition
            IAutoFilterCondition firstCondition = filter.FirstCondition;
            firstCondition.ConditionOperator = ExcelFilterCondition.DoesNotContain;
            firstCondition.String = "1000.00";

            //Saving the workbook
            workbook.SaveAs(Path.GetFullPath("../../../Output/Output.xlsx"));
        }
    }
}