using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Web.Mvc;
using Syncfusion.XlsIO;

namespace ExcelDataExport.Controllers
{
    public class HomeController : Controller
    {
        /// <summary>
        /// Displays the home page with the Export to XLSB button
        /// </summary>
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Handles the export request: reads from SQL Server, loads into DataTable,
        /// imports into Excel worksheet, and exports as XLSB
        /// </summary>
        public ActionResult ExportToXlsb()
        {
            try
            {
                // Step 1: Read data from SQL Server using ADO.NET
                DataTable productsTable = GetProductsFromDatabase();

                if (productsTable == null || productsTable.Rows.Count == 0)
                {
                    ViewBag.Message = "No data found in the database.";
                    return View("Index");
                }

                // Step 2: Create Excel workbook and worksheet
                using (IExcelEngine excelEngine = new ExcelEngine())
                {
                    IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);
                    IWorksheet worksheet = workbook.Worksheets[0];
                    worksheet.Name = "Products";

                    // Step 3: Import DataTable into worksheet
                    ImportDataTableToWorksheet(worksheet, productsTable);

                    // Step 4: Apply formatting (optional)
                    FormatWorksheet(worksheet);

                    // Step 5: Save as XLSB format using native Syncfusion API
                    byte[] excelBytes = workbook.SaveAsStream().ToArray();

                    // Step 6: Return file for download
                    return File(excelBytes, 
                                "application/vnd.ms-excel.sheet.binary.macroEnabled", 
                                "ProductsExport.xlsb");
                }
            }
            catch (SqlException sqlEx)
            {
                ViewBag.Message = $"Database error: {sqlEx.Message}";
                return View("Index");
            }
            catch (Exception ex)
            {
                ViewBag.Message = $"Export error: {ex.Message}";
                return View("Index");
            }
        }

        /// <summary>
        /// Reads product data from SQL Server database
        /// </summary>
        private DataTable GetProductsFromDatabase()
        {
            DataTable dataTable = new DataTable("Products");

            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    string query = @"
                        SELECT 
                            ProductID,
                            ProductName,
                            Category,
                            UnitPrice,
                            UnitsInStock,
                            DateAdded
                        FROM Products
                        ORDER BY ProductID";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.CommandTimeout = 30;

                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                    }
                }
                catch (SqlException ex)
                {
                    throw new Exception($"Failed to retrieve data from database: {ex.Message}");
                }
            }

            return dataTable;
        }

        /// <summary>
        /// Imports DataTable into Excel worksheet
        /// </summary>
        private void ImportDataTableToWorksheet(IWorksheet worksheet, DataTable dataTable)
        {
            // Import data into worksheet starting from A1
            worksheet.ImportDataTable(dataTable, true, 1, 1);
        }

        /// <summary>
        /// Applies basic formatting to the worksheet
        /// </summary>
        private void FormatWorksheet(IWorksheet worksheet)
        {
            // Format header row
            IRange headerRange = worksheet.Range[1, 1, 1, worksheet.UsedRange.LastColumn];
            headerRange.CellStyle.Font.Bold = true;
            headerRange.CellStyle.Font.Color = ExcelKnownColors.White;
            headerRange.CellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);

            // Auto-fit columns
            for (int i = 1; i <= worksheet.UsedRange.LastColumn; i++)
            {
                worksheet.AutofitColumn(i);
            }

            // Set minimum column width
            for (int i = 1; i <= worksheet.UsedRange.LastColumn; i++)
            {
                if (worksheet.GetColumnWidth(i) < 10)
                {
                    worksheet.SetColumnWidth(i, 12);
                }
            }
        }
    }
}
