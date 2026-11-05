using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Web.Mvc;
using ExcelDataExport.Data;
using Syncfusion.XlsIO;

namespace ExcelDataExport.Controllers
{
    public class HomeController : Controller
    {
        private const string XlsbContentType = "application/vnd.ms-excel.sheet.binary.macroEnabled.12";

        // GET: /
        public ActionResult Index()
        {
            return View();
        }

        // POST: /Home/ExportToXlsb
        // Access (.mdb) -> DataTable -> Excel worksheet -> XLSB download
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult ExportToXlsb()
        {
            try
            {
                // Step 1: Read data from the Access .mdb file into a DataTable
                DataTable productsTable = AccessDatabase.GetProducts();

                if (productsTable.Rows.Count == 0)
                {
                    ViewBag.Message = "No data found in the Products table.";
                    return View("Index");
                }

                // Step 2: Create the workbook and import the DataTable
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    IWorkbook workbook = application.Workbooks.Create(1);
                    workbook.Version = ExcelVersion.Xlsx;
                    IWorksheet worksheet = workbook.Worksheets[0];
                    worksheet.Name = "Products";

                    // true = first row is the header (column names), start at A1
                    worksheet.ImportDataTable(productsTable, true, 1, 1);

                    FormatWorksheet(worksheet, productsTable.Rows.Count);

                    // Step 3: Save as a real XLSB file using SaveAs(Stream, ExcelSaveType)
                    using (MemoryStream stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream, ExcelSaveType.SaveAsXLSB);
                        return File(stream.ToArray(), XlsbContentType, "ProductsExport.xlsb");
                    }
                }
            }
            catch (OleDbException ex)
            {
                ViewBag.Message = "Database error: " + ex.Message;
                return View("Index");
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Export error: " + ex.Message;
                return View("Index");
            }
        }


        /// <summary>
        /// Header styling, number/date formats and column widths.
        /// Columns: A ProductID, B ProductName, C Category, D UnitPrice, E UnitsInStock, F DateAdded
        /// </summary>
        private static void FormatWorksheet(IWorksheet worksheet, int rowCount)
        {
            int lastColumn = worksheet.UsedRange.LastColumn;
            int lastRow = rowCount + 1; // +1 for the header row

            IRange header = worksheet.Range[1, 1, 1, lastColumn];
            header.CellStyle.Font.Bold = true;
            header.CellStyle.Font.Color = ExcelKnownColors.White;
            header.CellStyle.Color = System.Drawing.Color.FromArgb(0, 102, 204);

            worksheet.Range[2, 4, lastRow, 4].NumberFormat = "#,##0.00";        // UnitPrice
            worksheet.Range[2, 6, lastRow, 6].NumberFormat = "yyyy-mm-dd hh:mm"; // DateAdded

            worksheet.UsedRange.AutofitColumns();
        }
    }
}
