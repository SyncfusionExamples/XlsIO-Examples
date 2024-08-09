using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Diagnostics;
using Syncfusion.XlsIO;
using TemplateMarker_with_Formulas.Models;

namespace TemplateMarker_with_Formulas.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private static DataTable northwindDt;
        private static DataTable numbersDt;
        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }
        public IActionResult Index()
        {
            return View();
        }
        private static IList<Customer> GetCustomerAsObjects()
        {
            DataSet customersDataSet = new DataSet();
            FileStream dataStream = new FileStream("Data/customers.xml", FileMode.Open, FileAccess.Read);
            dataStream.Position = 0;
            customersDataSet.ReadXml(dataStream, XmlReadMode.ReadSchema);
            northwindDt = customersDataSet.Tables[0];
            IList<Customer> tmpCustomers = new List<Customer>();
            Customer customer = new Customer();
            numbersDt = GetTable();
            DataRowCollection rows = northwindDt.Rows;
            foreach (DataRow row in rows)
            {
                customer = new Customer();
                customer.SalesPerson = row[0].ToString();
                customer.SalesJanJune = Convert.ToInt32(row[1]);
                customer.SalesJulyDec = Convert.ToInt32(row[2]);
                customer.Image = GetImage(Convert.ToString(row[4]));
                tmpCustomers.Add(customer);
            }
            return tmpCustomers;
        }
        private static byte[] GetImage(string path)
        {
            FileStream imageStream = new FileStream("Images/" + path, FileMode.Open, FileAccess.Read);
            using (BinaryReader reader = new BinaryReader(imageStream))
            {
                return reader.ReadBytes((int)imageStream.Length);
            }
        }
        private static DataTable GetTable()
        {
            Random r = new Random();
            DataTable dt = new DataTable("NumbersTable");

            int nCols = 4;
            int nRows = 10;

            for (int i = 0; i < nCols; i++)
                dt.Columns.Add(new DataColumn("Column" + i.ToString()));

            for (int i = 0; i < nRows; ++i)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < nCols; j++)
                    dr[j] = r.Next(0, 10);
                dt.Rows.Add(dr);
            }
            return dt;
        }
        public IActionResult TemplateMarker_with_Formulas()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //load an existing file
                FileStream excelStream = new FileStream("Data/Formulas.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(excelStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Create Template Marker Processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                //Add marker variable
                marker.AddVariable("NumbersTable", GetTable());

                //Process the markers in the template
                marker.ApplyMarkers();

                worksheet.Activate();

                //Save and close the workbook
                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "Output.xlsx";
                workbook.Close();
                excelEngine.Dispose();

                return fileStreamResult;
            }
        }
        public IActionResult Privacy()
        {
            return View();
        }
        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}