using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Read_Custom_XML
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //Access CustomXmlPart from Workbook
                ICustomXmlPart customXmlPart = workbook.CustomXmlparts.GetById("SD10003");

                //Access XmlData from CustomXmlPart
                byte[] xmlData = customXmlPart.Data;

                System.Text.Encoding.Default.GetString(xmlData);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ReadXml.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}





