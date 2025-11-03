using System;
using Syncfusion.XlsIO;
using System.IO;

namespace Add_Custom_XML
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding CustomXmlData to Workbook
                ICustomXmlPart customXmlPart = workbook.CustomXmlparts.Add("SD10003"); 

                //Add XmlData to CustomXmlPart
                byte[] xmlData = File.ReadAllBytes(Path.GetFullPath(@"Data/InputTemplate.xml"));
                customXmlPart.Data = xmlData;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/CreateCustomXML.xlsx"));
                #endregion
                //Open default JSON
            }
        }
    }
}