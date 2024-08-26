using Syncfusion.XlsIO;

namespace ExcelTable
{
    class Program
    {
        public static void Main()
        {
            Program program = new Program();
            program.CreateTable();
        }

        /// <summary>
        /// Creates a Excel table in the existing Excel document.
        /// </summary>
        public void CreateTable()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream fileStream = new FileStream("../../../Data/SalesReport.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream, ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];

                AddExcelTable("Table1", worksheet.UsedRange);
                
                string fileName = Path.GetFullPath(@"Output/SalesReport.xlsx");

                //Saving the workbook as stream
                FileStream stream = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);
                stream.Dispose();
            }
        }

        /// <summary>
        /// Adds a table to the worksheet with the given name and range
        /// </summary>
        /// <param name="tableName">Table name</param>
        /// <param name="tableRange">Table range</param>
        public void AddExcelTable(string tableName, IRange tableRange)
        {
            IWorksheet worksheet = tableRange.Worksheet;

            //Create table with the data in given range
            IListObject table = worksheet.ListObjects.Create(tableName, tableRange);

            //Set table style
            table.BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium14;
        }
    }
}





