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
                byte[] xmlData = File.ReadAllBytes("../../../Data/InputTemplate.xml");
                customXmlPart.Data = xmlData;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("CreateCustomXML.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                //Open default JSON
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("CreateCustomXML.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
