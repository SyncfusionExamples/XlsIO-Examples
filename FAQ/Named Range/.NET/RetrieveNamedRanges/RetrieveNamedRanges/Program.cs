using Syncfusion.XlsIO;

class Program
{        
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));
            IWorksheet worksheet = workbook.Worksheets[0];

            //Retrieving names defined in the workbook 
            IName[] names = new IName[workbook.Names.Count];
            for (int i = 0; i < workbook.Names.Count; i++)
            {
                names[i] = workbook.Names[i];
                Console.WriteLine(names[i].Name);
            }

            //Saving the workbook
            workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
        }
    }
}