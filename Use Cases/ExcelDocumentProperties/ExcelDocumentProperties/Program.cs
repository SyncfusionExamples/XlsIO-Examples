//Create an instance of ExcelEngine
using Syncfusion.Drawing;
using Syncfusion.XlsIO;
using System.Globalization;
using static System.Net.Mime.MediaTypeNames;

namespace ExcelDocumentProperties
{

    class Program
    {
        static void Main(string[] args)
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Create a workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Generate Invoice
                AddInvoiceDetails(worksheet);

                //Apply built-in and custom document properties
                ApplyDocumentProperties(workbook);

                workbook.SaveAs(Path.GetFullPath("DocumentProperties.xlsx"));
            }
        }
        /// <summary>
        /// Apply built-in and custom document properties to the workbook
        /// </summary>
        /// <param name="workbook">IWorkbook</param>
        public static void ApplyDocumentProperties(IWorkbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            // Read key invoice details from the worksheet
            int invoiceNumber = (int)worksheet.Range["D6"].Number;         
            string invoiceDateText = worksheet.Range["E6"].DisplayText;    
            int customerId = (int)worksheet.Range["D8"].Number;            
            string terms = worksheet.Range["E8"].DisplayText;              
            string customerName = worksheet.Range["A8"].DisplayText;       
            string customerCompany = worksheet.Range["A9"].DisplayText;    
            DateTime invoiceDate = DateTime.Now;


            // Add the document properties for the invoice 
            IBuiltInDocumentProperties builtInProperties = workbook.BuiltInDocumentProperties;
            builtInProperties.Title = $"Invoice #{invoiceNumber}";
            builtInProperties.Author = "Jim Halper";
            builtInProperties.Subject = $"Invoice for {customerName} ({customerCompany})";
            builtInProperties.Keywords = $"invoice;billing;customer:{customerId};terms:{terms}";
            builtInProperties.Company = "Great Lake Enterprises";
            builtInProperties.Category = "Finance/Billing";
            builtInProperties.Comments = $"Issued {invoiceDate:yyyy-MM-dd}";

            // Add the custom document properties for the invoice
            var customProperties = workbook.CustomDocumentProperties;
            customProperties["InvoiceNumber"].Value = invoiceNumber;
            customProperties["InvoiceDate"].Text = invoiceDate.ToString("yyyy-MM-dd");
            customProperties["CustomerId"].Value = customerId;
            customProperties["CustomerName"].Text = customerName;
            customProperties["CustomerCompany"].Text = customerCompany;
            customProperties["Currency"].Text = "USD";
            customProperties["PaymentStatus"].Text = "Completed";
            customProperties["Confidential"].Value = true;
        }

        /// <summary>
        /// Populates the Invoice details to the worksheet
        /// </summary>
        /// <param name="worksheet">Worksheet</param>
        public static void AddInvoiceDetails(IWorksheet worksheet)
        {
            //Disable gridlines in the worksheet
            worksheet.IsGridLinesVisible = false;

            //Enter text to the cell A1 and apply formatting.
            worksheet.Range["A1:D1"].Merge();
            worksheet.Range["A1"].Text = "SALES INVOICE";
            worksheet.Range["A1"].CellStyle.Font.Bold = true;
            worksheet.Range["A1"].CellStyle.Font.RGBColor = Color.FromArgb(42, 118, 189);
            worksheet.Range["A1"].CellStyle.Font.Size = 35;

            //Apply alignment in the cell D1
            worksheet.Range["D1"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;
            worksheet.Range["D1"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;

            //Enter values to the cells from A2 to A5
            worksheet.Range["A2"].Text = "Great Lake Bike Parts";
            worksheet.Range["A3"].Text = "46036 Michigan Ave";
            worksheet.Range["A4"].Text = "Canton, USA";
            worksheet.Range["A5"].Text = "Phone: +1 231-231-2310";

            //Make the text bold
            worksheet.Range["A2:A5"].CellStyle.Font.Bold = true;

            //Merge cells
            worksheet.Range["D1:E1"].Merge();

            //Enter values to the cells from D5 to E8
            worksheet.Range["D5"].Text = "INVOICE#";
            worksheet.Range["E5"].Text = "DATE";
            worksheet.Range["D6"].Number = 1028;
            worksheet.Range["E6"].Value = DateTime.Now.ToString("yyyy-MM-dd");
            worksheet.Range["D7"].Text = "CUSTOMER ID";
            worksheet.Range["E7"].Text = "Payment Status";
            worksheet.Range["D8"].Number = 564;
            worksheet.Range["E8"].Text = "Completed";

            //Apply RGB backcolor to the cells from D5 to E8
            worksheet.Range["D5:E5"].CellStyle.Color = Color.FromArgb(42, 118, 189);
            worksheet.Range["D7:E7"].CellStyle.Color = Color.FromArgb(42, 118, 189);

            //Apply known colors to the text in cells D5 to E8
            worksheet.Range["D5:E5"].CellStyle.Font.Color = ExcelKnownColors.White;
            worksheet.Range["D7:E7"].CellStyle.Font.Color = ExcelKnownColors.White;

            //Make the text as bold from D5 to E8
            worksheet.Range["D5:E8"].CellStyle.Font.Bold = true;

            //Apply alignment to the cells from D5 to E8
            worksheet.Range["D5:E8"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            worksheet.Range["D5:E5"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
            worksheet.Range["D7:E7"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
            worksheet.Range["D6:E6"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;

            //Enter value and applying formatting in the cell A7
            worksheet.Range["A7"].Text = "  BILL TO";
            worksheet.Range["A7"].CellStyle.Color = Color.FromArgb(42, 118, 189);
            worksheet.Range["A7"].CellStyle.Font.Bold = true;
            worksheet.Range["A7"].CellStyle.Font.Color = ExcelKnownColors.White;

            //Apply alignment
            worksheet.Range["A7"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
            worksheet.Range["A7"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

            //Enter values in the cells A8 to A12
            worksheet.Range["A8"].Text = "Steyn";
            worksheet.Range["A9"].Text = "20 Whitehall Rd";
            worksheet.Range["A10"].Text = "North Muskegon,USA";
            worksheet.Range["A11"].Text = "+1 231-654-0000";

            //Create a Hyperlink for e-mail in the cell A13
            IHyperLink hyperlink = worksheet.HyperLinks.Add(worksheet.Range["A12"]);
            hyperlink.Type = ExcelHyperLinkType.Url;
            hyperlink.Address = "steyn@xyz.com";
            hyperlink.ScreenTip = "Send Mail";

            //Merge column A and B from row 15 to 22
            worksheet.Range["A15:B15"].Merge();
            worksheet.Range["A16:B16"].Merge();
            worksheet.Range["A17:B17"].Merge();
            worksheet.Range["A18:B18"].Merge();
            worksheet.Range["A19:B19"].Merge();
            worksheet.Range["A20:B20"].Merge();
            worksheet.Range["A21:B21"].Merge();
            worksheet.Range["A22:B22"].Merge();

            // Headers
            worksheet.Range["A15"].Text = " Items";
            worksheet.Range["C15"].Text = "QTY";
            worksheet.Range["D15"].Text = "UNIT PRICE";
            worksheet.Range["E15"].Text = "AMOUNT";

            // Bike spare parts
            worksheet.Range["A16"].Text = "Brake Pads";
            worksheet.Range["A17"].Text = "Chain";
            worksheet.Range["A18"].Text = "Gear Cable Set";
            worksheet.Range["A19"].Text = "Pedals";
            worksheet.Range["A20"].Text = "Tyre (700x25C)";

            // Quantities
            worksheet.Range["C16"].Number = 2;  
            worksheet.Range["C17"].Number = 1;  
            worksheet.Range["C18"].Number = 2;  
            worksheet.Range["C19"].Number = 1;  
            worksheet.Range["C20"].Number = 2;  

            // Unit Prices (in USD)
            worksheet.Range["D16"].Number = 15;  
            worksheet.Range["D17"].Number = 35;  
            worksheet.Range["D18"].Number = 12;  
            worksheet.Range["D19"].Number = 45;
            worksheet.Range["D20"].Number = 30;

            worksheet.Range["D23"].Text = "Total";

            //Apply number format
            worksheet.Range["D16:E22"].NumberFormat = "$0.00";
            worksheet.Range["E23"].NumberFormat = "$0.00";

            //Apply incremental formula for column Amount by multiplying Qty and UnitPrice
            worksheet.Application.EnableIncrementalFormula = true;
            worksheet.Range["E16:E20"].Formula = "=C16*D16";

            //Formula for Sum the total
            worksheet.Range["E23"].Formula = "=SUM(E16:E22)";

            //Apply borders
            worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
            worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
            worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Grey_25_percent;
            worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Grey_25_percent;
            worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
            worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
            worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Black;
            worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Black;

            //Apply font setting for cells with product details
            worksheet.Range["A3:E23"].CellStyle.Font.FontName = "Arial";
            worksheet.Range["A3:E23"].CellStyle.Font.Size = 10;
            worksheet.Range["A15:E15"].CellStyle.Font.Color = ExcelKnownColors.White;
            worksheet.Range["A15:E15"].CellStyle.Font.Bold = true;
            worksheet.Range["D23:E23"].CellStyle.Font.Bold = true;

            //Apply cell color
            worksheet.Range["A15:E15"].CellStyle.Color = Color.FromArgb(42, 118, 189);

            //Apply alignment to cells with product details
            worksheet.Range["A15"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
            worksheet.Range["C15:C22"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            worksheet.Range["D15:E15"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;

            //Apply row height and column width to look good
            worksheet.Range["A1"].ColumnWidth = 36;
            worksheet.Range["B1"].ColumnWidth = 11;
            worksheet.Range["C1"].ColumnWidth = 8;
            worksheet.Range["D1:E1"].ColumnWidth = 18;
            worksheet.Range["A1"].RowHeight = 47;
            worksheet.Range["A2"].RowHeight = 15;
            worksheet.Range["A3:A4"].RowHeight = 15;
            worksheet.Range["A5"].RowHeight = 18;
            worksheet.Range["A6"].RowHeight = 29;
            worksheet.Range["A7"].RowHeight = 18;
            worksheet.Range["A8"].RowHeight = 15;
            worksheet.Range["A9:A14"].RowHeight = 15;
            worksheet.Range["A15:A23"].RowHeight = 18;
        }
    }
}