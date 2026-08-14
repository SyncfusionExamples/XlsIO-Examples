using Syncfusion.XlsIO;
using Syncfusion.Compression.Zip;

namespace SplitExcel
{
    class Program
    {
        private static readonly string inputPath = Path.GetFullPath("Data/");
        private static readonly string outputPath = Path.GetFullPath("Output/");

        static void Main(string[] args)
        {
            string fileName = "Report.xlsx";
            SplitExcelDocument(Path.GetFullPath(inputPath + fileName));
        }

        /// <summary>
        /// Split the Excel document from the given path and add outputs into a ZIP archive.
        /// </summary>
        private static void SplitExcelDocument(string filePath)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(filePath);
                IWorksheets worksheets = workbook.Worksheets;

                workbook.Version = ExcelVersion.Xlsx;

                // Create a ZIP archive with best compression
                ZipArchive zipArchive = new ZipArchive();
                zipArchive.DefaultCompressionLevel = Syncfusion.Compression.CompressionLevel.Best;

                foreach (IWorksheet worksheet in worksheets)
                {
                    IWorkbook newBook = application.Workbooks.Create(0);
                    newBook.Worksheets.AddCopy(worksheet);

                    string sheetFilePath = Path.GetFullPath(worksheet.Name + ".xlsx");

                    newBook.SaveAs(sheetFilePath);

                    // Add that file directly into the ZIP archive
                    zipArchive.AddFile(sheetFilePath);

                    newBook.Close();
                }

                // Save the ZIP archive containing all worksheets
                string zipFilePath = Path.GetFullPath(outputPath + "Split-Excel-Sheets.zip");
                zipArchive.Save(zipFilePath);
                zipArchive.Close();

                workbook.Close();
            }
        }
    }
}