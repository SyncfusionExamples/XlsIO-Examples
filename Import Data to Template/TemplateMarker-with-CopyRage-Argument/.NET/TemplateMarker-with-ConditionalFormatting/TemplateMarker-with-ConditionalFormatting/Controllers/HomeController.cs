using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;
using TemplateMarker_with_ConditionalFormatting.Models;
using System.Data;

namespace TemplateMarker_with_ConditionalFormatting.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private static DataTable northwindDt;
        private static DataTable numbersDt;
        public static IList<Customer> _customers = new List<Customer>();
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
        public IActionResult TemplateMarker_with_ConditionalFormatting()
        {

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //load an existing file
                FileStream excelStream = new FileStream("Data/TemplateMarkerImages.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(excelStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Create Template Marker Processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                //Applying conditional formats

                #region Data Bar
                IConditionalFormats conditions = marker.CreateConditionalFormats(worksheet["C5"]);
                IConditionalFormat condition = conditions.AddCondition();

                //Set Data bar and icon set for the same cell
                //Set the format type
                condition.FormatType = ExcelCFType.DataBar;
                IDataBar dataBar = condition.DataBar;

                //Set the constraint
                dataBar.MinPoint.Type = ConditionValueType.LowestValue;
                dataBar.MaxPoint.Type = ConditionValueType.HighestValue;

                //Set color for Bar
                dataBar.BarColor = Color.FromArgb(156, 208, 243);

                //Hide the value in data bar
                dataBar.ShowValue = false;
                #endregion

                #region Color Scale
                conditions = marker.CreateConditionalFormats(worksheet["D5"]);
                condition = conditions.AddCondition();

                condition.FormatType = ExcelCFType.ColorScale;
                IColorScale colorScale = condition.ColorScale;

                //Sets 3 - color scale
                colorScale.SetConditionCount(3);

                colorScale.Criteria[1].FormatColorRGB = Color.FromArgb(244, 210, 178);
                colorScale.Criteria[1].Type = ConditionValueType.Percentile;
                colorScale.Criteria[1].Value = "50";

                colorScale.Criteria[2].FormatColorRGB = Color.FromArgb(245, 247, 171);
                colorScale.Criteria[2].Type = ConditionValueType.Percentile;
                colorScale.Criteria[2].Value = "100";
                #endregion

                //Add marker variable
                marker.AddVariable("SalesList", GetCustomerAsObjects());

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