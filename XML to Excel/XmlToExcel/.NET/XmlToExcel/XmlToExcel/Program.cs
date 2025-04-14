using Syncfusion.XlsIO;

namespace XmlToExcel
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Import Xml data into the worksheet
                FileStream inputStream = new FileStream("../../../Data/XmlFile.xml", FileMode.Open, FileAccess.Read);

                worksheet.ImportXml(inputStream, 1, 1);

                worksheet.UsedRange.AutofitColumns();

                //Saving the workbook as stream
                FileStream outputStream = new FileStream("../../../Output/Output.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(outputStream);

                //Dispose stream
                inputStream.Dispose();
                outputStream.Dispose();
            }
        }
    }
}
