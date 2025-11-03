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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //Access CustomXmlPart from Workbook
                ICustomXmlPart customXmlPart = workbook.CustomXmlparts.GetById("SD10003");

                //Access XmlData from CustomXmlPart
                byte[] xmlData = customXmlPart.Data;

                System.Text.Encoding.Default.GetString(xmlData);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ReadXml.xlsx"));
                #endregion
            }
        }
    }
}