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
            FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Input.xlsx"), FileMode.Open, FileAccess.Read);
            IWorkbook workbook = application.Workbooks.Open(inputStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Creating an AutoFilter 
            worksheet.AutoFilters.FilterRange = worksheet.Range["A1:A11"];
            IAutoFilter filter = worksheet.AutoFilters[0];

            //Specifying first condition
            IAutoFilterCondition firstCondition = filter.FirstCondition;
            firstCondition.ConditionOperator = ExcelFilterCondition.DoesNotContain;
            firstCondition.String = "1000.00";

            //Saving the workbook
            FileStream outputStream = new FileStream(Path.GetFullPath("../../../Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
            workbook.SaveAs(outputStream);

            //Dispose streams
            outputStream.Dispose();
            inputStream.Dispose();
        }
    }
}