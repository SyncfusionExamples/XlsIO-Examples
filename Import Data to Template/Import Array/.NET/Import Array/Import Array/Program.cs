using Syncfusion.XlsIO;

namespace Import_Array
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Initialize Excel engine and application.
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                // Open an existing workbook.
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);

                // Create Template Marker Processor.
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                // Insert Array Horizontally.
                string[] names = { "Mickey", "Donald", "Tom", "Jerry" };
                string[] descriptions = { "Mouse", "Duck", "Cat", "Mouse" };

                // Add collections to the marker variables where the name should match with input template.
                marker.AddVariable("Names", names);
                marker.AddVariable("Descriptions", descriptions);

                // Process the markers in the template.
                marker.ApplyMarkers();

                // Saving the workbook.
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ImportArray.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                inputStream.Dispose();
                outputStream.Dispose();
            }
        }
    }
}