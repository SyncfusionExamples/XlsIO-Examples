using Syncfusion.XlsIO;

namespace MergeExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = Path.GetFullPath("Data");
            string outputPath = Path.GetFullPath("Output");

            FileInfo[] files = new DirectoryInfo(inputPath).GetFiles();

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook mergedBook = application.Workbooks.Create(0);

                foreach (FileInfo file in files)
                {
                    IWorkbook tempWorkbook = application.Workbooks.Open(Path.GetFullPath(file.FullName));
                    mergedBook.Worksheets.AddCopy(tempWorkbook.Worksheets);
                    tempWorkbook.Close();
                }

                mergedBook.Version = ExcelVersion.Xlsx;
                mergedBook.SaveAs(Path.GetFullPath(Path.Combine(outputPath, "MergedExcel.xlsx")));
            }
        }
    }
}