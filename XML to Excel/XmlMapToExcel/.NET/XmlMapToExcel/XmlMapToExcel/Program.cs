using Syncfusion.XlsIO;

namespace XmlMapToExcel
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

                //Import XML mapping to Excel
                workbook.XmlMaps.Add(inputStream);

                //Saving the workbook as stream
                FileStream outputStream = new FileStream("../../../Output/XmlMapOutput.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(outputStream);

                //Dispose stream
                inputStream.Dispose();
                outputStream.Dispose();
            }
        }
    }
}

