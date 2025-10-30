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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);

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
                workbook.SaveAs(Path.GetFullPath("Output/ImportArray.xlsx"));

            }
        }
    }
}