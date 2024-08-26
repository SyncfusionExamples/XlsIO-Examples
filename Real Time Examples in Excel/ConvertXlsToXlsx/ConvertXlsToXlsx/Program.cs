using Syncfusion.XlsIO;


namespace ConvertXlsToXlsx
{
    class Program
    {
        private static string inputPath = @"../../../Data/";

        private static string outputPath = @"../../../Output/";
        static void Main(string[] args)
        {
            string fileName = "Report.xls";

            //Split the Excel document
            ConvertXlsToXLSX(inputPath + fileName);
        }
        /// <summary>
        /// Convert the Excel document from the given path
        /// </summary>
        /// <param name="filePath">Excel file path</param>
        private static void ConvertXlsToXLSX(string filePath)
        {
            FileStream inputData = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite);

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(inputData);
                IWorksheets worksheets = workbook.Worksheets;

                workbook.Version = ExcelVersion.Xlsx;

                FileStream fileStream = new FileStream(outputPath + Path.GetFileNameWithoutExtension(filePath) + ".xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(fileStream);
                fileStream.Close();
            }
        }
    }
}






