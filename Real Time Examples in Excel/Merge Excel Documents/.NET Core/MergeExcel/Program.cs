// See https://aka.ms/new-console-template for more information
using Syncfusion.XlsIO;


namespace MergeExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = @"../../../Data/";

            string outputPath = @"../../../Output/";

            FileInfo[] files = new DirectoryInfo(inputPath).GetFiles();

            List<Stream> streams = new List<Stream>();

            foreach (FileInfo file in files)
            {
                streams.Add(file.OpenRead());
            }

            Stream mergedStream = MergeExcelDocuments(streams);

            FileStream fileStream = new FileStream(outputPath + "MergedExcel.xlsx", FileMode.Create, FileAccess.Write);
            mergedStream.Position = 0;
            mergedStream.CopyTo(fileStream);
            fileStream.Close();
        }

        /// <summary>
        /// Merge Excel documents from the list of Excel streams
        /// </summary>
        /// <param name="streams">List of Excel document stream to be merged</param>
        /// <returns></returns>
        public static Stream MergeExcelDocuments(List<Stream> streams)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(0);

                //Loop through each Excel document and add the worksheets to the new workbook
                foreach (Stream stream in streams)
                {
                    stream.Position = 0;
                    IWorkbook tempWorkbook = application.Workbooks.Open(stream);
                    workbook.Worksheets.AddCopy(tempWorkbook.Worksheets);
                    tempWorkbook.Close();
                }

                //Save the workbook to a memory stream
                MemoryStream memoryStream = new MemoryStream();
                workbook.Version = ExcelVersion.Xlsx;
                workbook.SaveAs(memoryStream);

                return memoryStream;
            }
        }
    }
}


