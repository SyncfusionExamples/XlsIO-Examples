using Syncfusion.XlsIO;


namespace SplitExcel
{
    class Program
    {
        private static string inputPath = @"../../../Data/";

        private static string outputPath = @"../../../Output/";
        static void Main(string[] args)
        {
            string fileName = "Report.xlsx";

            //Split the Excel document
            SplitExcelDocument(Path.GetFullPath(inputPath + fileName));
        }
        /// <summary>
        /// Split the Excel document from the given path
        /// </summary>
        /// <param name="filePath">Excel file path</param>
        private static void SplitExcelDocument(string filePath)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(filePath);
                IWorksheets worksheets = workbook.Worksheets;

                workbook.Version = ExcelVersion.Xlsx;

                //Loop through each Excel worksheet and save it as a new workbook
                foreach (IWorksheet worksheet in worksheets)
                {
                    IWorkbook newBook = application.Workbooks.Create(0);
                    newBook.Worksheets.AddCopy(worksheet);

                    newBook.SaveAs(Path.GetFullPath(outputPath + worksheet.Name + ".xlsx"));
                }
            }
        }
    }
}






