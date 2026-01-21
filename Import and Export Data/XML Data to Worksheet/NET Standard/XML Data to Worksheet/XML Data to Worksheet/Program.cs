using Syncfusion.XlsIO;

namespace XML_Data_to_Worksheet
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
                worksheet.ImportXml(@"../../../Data/XmlFile.xml", 1, 6);

                //Saving the workbook as stream
                FileStream stream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);

                //Dispose stream
                stream.Dispose();
            }
        }
    }
}
