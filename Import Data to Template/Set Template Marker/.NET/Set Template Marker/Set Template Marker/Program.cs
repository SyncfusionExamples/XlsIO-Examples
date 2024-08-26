using System;
using System.IO;
using Syncfusion.XlsIO;

namespace Set_Template_Marker
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
                IWorksheet worksheet = workbook.Worksheets[0];

                //Insert Simple marker
                worksheet.Range["B2"].Text = "%Marker";

                //Insert marker which gets value of Author name
                worksheet.Range["C2"].Text = "%Marker2.Worksheet.Workbook.Author";

                //Create Template Marker Processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                //Add collections to the marker variables where the name should match with input template
                marker.AddVariable("Marker", new DateTime(2017, 03, 02));
                marker.AddVariable("Marker2", worksheet.Range["B2"]);

                //Process the markers in the template
                marker.ApplyMarkers();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/TemplateMarker.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}

