using Syncfusion.XlsIO;
using static System.Net.Mime.MediaTypeNames;

using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Open(@"..\..\..\Data\Input.xlsx", ExcelOpenType.Automatic);
    IWorksheet worksheet = workbook.Worksheets[0];

    ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

    List<string> fruits = new List<string>();
    fruits.Add("Apple");
    fruits.Add("Banana");
    fruits.Add("Orange");
    fruits.Add("Mango");
    fruits.Add("Blueberry");
    fruits.Add("Pineapple");

    List<string> places = new List<string>();
    places.Add("New York");
    places.Add("London");
    places.Add("Tokyo");
    places.Add("Paris");


    List<string> cars = new List<string>();
    cars.Add("Toyota Corolla");
    cars.Add("Honda Civic");


    marker.AddVariable("Places", places);
    marker.AddVariable("Fruits", fruits);
    marker.AddVariable("Cars", cars);
    marker.ApplyMarkers();


    workbook.SaveAs(@"..\..\..\Data\Output.xlsx");
    workbook.Close();

}