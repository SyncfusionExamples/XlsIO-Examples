using Syncfusion.XlsIO;

using Syncfusion.XlsIO.Implementation;
using System.Reflection;

class Program
{
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            //Instantiate the Excel application object
            IApplication application = excelEngine.Excel;

            //Assigns default application version
            application.DefaultVersion = ExcelVersion.Xlsx;

            //A new workbook is created equivalent to creating a new workbook in Excel
            //Create a workbook with 1 worksheet
            IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));

            //Access first worksheet from the workbook
            IWorksheet worksheet = workbook.Worksheets[0];

            IConditionalFormats conditionalFormats = worksheet["B3"].ConditionalFormats;
            List<ConditionalFormatImpl> conditionalformats = GetSortedConditionalFormats(conditionalFormats);

            for (int i = 0; i < conditionalformats.Count; i++)
            {
                IConditionalFormat format = conditionalformats[i];
                Console.WriteLine(format.FirstFormula);
            }

            //Saving the workbook 
            workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
        }
    }
    static List<ConditionalFormatImpl> GetSortedConditionalFormats(IConditionalFormats conditionalFormats)
    {
        List<ConditionalFormatImpl> result = new List<ConditionalFormatImpl>();

        //Reflection
        MethodInfo getConditionMethod = typeof(ConditionalFormatWrapper).GetMethod("GetCondition", BindingFlags.Instance | BindingFlags.NonPublic);
        PropertyInfo priorityProp = typeof(ConditionalFormatImpl).GetProperty("Priority", BindingFlags.Instance | BindingFlags.NonPublic);

        for (int i = 0; i < conditionalFormats.Count; i++)
        {
            IConditionalFormat format = conditionalFormats[i];
            ConditionalFormatImpl impl = format as ConditionalFormatImpl;

            if (impl == null && format is ConditionalFormatWrapper wrapper)
            {
                impl = getConditionMethod.Invoke(wrapper, null) as ConditionalFormatImpl;
            }
            if (impl != null)
            {
                result.Add(impl);
            }
        }
        return result.OrderBy(f => (int)priorityProp.GetValue(f)).ToList();
    }
}