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
            SplitExcelDocument(inputPath + fileName);
        }
        /// <summary>
        /// Split the Excel document from the given path
        /// </summary>
        /// <param name="filePath">Excel file path</param>
        private static void SplitExcelDocument(string filePath)
        {
            FileStream inputData = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite);

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(inputData);
                IWorksheets worksheets = workbook.Worksheets;

                workbook.Version = ExcelVersion.Xlsx;

                //Loop through each Excel worksheet and save it as a new workbook
                foreach (IWorksheet worksheet in worksheets)
                {
                    IWorkbook newBook = application.Workbooks.Create(0);
                    newBook.Worksheets.AddCopy(worksheet);

                    FileStream outputData = new FileStream(outputPath + worksheet.Name + ".xlsx", FileMode.Create, FileAccess.ReadWrite);
                    newBook.SaveAs(outputData);
                    outputData.Close();
                }
            }
        }
    }
}

