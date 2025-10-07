using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using System;
using static System.Net.Mime.MediaTypeNames;

class Program
{
    static void Main()
    {
        using (ExcelEngine engine = new ExcelEngine())
        {
            IApplication app = engine.Excel;
            app.DefaultVersion = ExcelVersion.Xlsx;

            IWorkbook workbook = app.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Monthly Sales";

            // Title
            worksheet.Range["A1"].Text = "Monthly Sales Report";
            worksheet.Range["A1"].CellStyle.Font.Bold = true;
            worksheet.Range["A1"].CellStyle.Font.Size = 16;

            // Headers
            worksheet.Range["A3"].Text = "Order ID";
            worksheet.Range["B3"].Text = "Date";
            worksheet.Range["C3"].Text = "Region";
            worksheet.Range["D3"].Text = "Salesperson";
            worksheet.Range["E3"].Text = "Units";
            worksheet.Range["F3"].Text = "Amount";
            worksheet.Range["A3:F3"].CellStyle.Font.Bold = true;
            worksheet.Range["A3:F3"].CellStyle.Color = Syncfusion.Drawing.Color.FromArgb(240, 240, 240);

            // Generate Sales Data
            int row = 4;
            string[] regions = new[] { "North", "South", "West", "East" };
            Random rnd = new Random(17);
            DateTime month = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);

            foreach (string region in regions)
            {
                worksheet.Range["A" + row].Text = region + " Region";
                worksheet.Range["A" + row].CellStyle.Font.Bold = true;
                worksheet.Range["A" + row].CellStyle.Font.Size = 12;
                row++;

                for (int count = 0; count < 20; count++)
                {
                    worksheet.Range["A" + row].Text = $"ORD-{region.Substring(0, 1)}-{1000 + count}";
                    worksheet.Range["B" + row].DateTime = month.AddDays(rnd.Next(0, 28));
                    worksheet.Range["B" + row].NumberFormat = "dd-MMM";
                    worksheet.Range["C" + row].Text = region;
                    worksheet.Range["D" + row].Text = "Rep " + rnd.Next(1, 6);
                    worksheet.Range["E" + row].Number = rnd.Next(1, 25);
                    worksheet.Range["F" + row].Number = Math.Round(2500 + rnd.NextDouble() * 20000, 2);
                    worksheet.Range["F" + row].NumberFormat = "#,##0.00";
                    row++;
                }
                row++;
            }

            worksheet.UsedRange.AutofitColumns();

            // Page Setup Options
            IPageSetup pageSetup = worksheet.PageSetup;

            // Set Paper Size to A4
            pageSetup.PaperSize = ExcelPaperSize.PaperA4;

            // Set Orientation to Landscape
            pageSetup.Orientation = ExcelPageOrientation.Landscape;

            // Set Margins in inches

            // Top, Bottom - 0.5 inch
            pageSetup.TopMargin = 0.75;
            pageSetup.BottomMargin = 0.75;

            // Left, Right - 0.25 inch
            pageSetup.LeftMargin = 0.25;
            pageSetup.RightMargin = 0.25;

            // Header, Footer - 0.3 inch
            pageSetup.HeaderMargin = 0.3;
            pageSetup.FooterMargin = 0.3;

            // Set fit to page as false. 
            pageSetup.IsFitToPage = false;

            // Set repeat rows as 3rd row
            pageSetup.PrintTitleRows = "$3:$3";

            // Define the print area from A3 to last used row in column F
            int lastRow = worksheet.UsedRange.LastRow;
            pageSetup.PrintArea = $"A3:F{lastRow}";

            // Apply left header as "Monthly Sales" with Calibri font of size 14 and bold
            pageSetup.LeftHeader = "&\"Calibri,Bold\"&14 Monthly Sales";

            // Apply center header as "Month Year" with bold
            pageSetup.CenterHeader = "&B&10" + month.ToString("MMMM yyyy");

            // Apply right header with page number and total pages
            pageSetup.RightHeader = "Page &P of &N";

            // Apply left footer with sheet name
            pageSetup.LeftFooter = "Sheet: &A";

            // Apply center footer with current date and time
            pageSetup.CenterFooter = "Generated: &D &T";

            // Apply right footer with file name
            pageSetup.RightFooter = "&F";

            // Set the scaling to 120%
            pageSetup.Zoom = 120;

            // Fit to 1 page wide and unlimited tall
            pageSetup.FitToPagesTall = 0;

            // Fit to 1 page wide
            pageSetup.FitToPagesWide = 1;

            // Add a page break before each region. Assuming each region block (title + data) takes 22 rows including spacing.
            int firstRegionRow = 4;
            int blockHeight = 22;
            for (int count = 1; count < regions.Length; count++)
            {
                int titleRow = firstRegionRow + count * blockHeight;

                // Add a horizontal page break before the title row of each region
                worksheet.HPageBreaks.Add(worksheet.Range["A" + titleRow]);
            }

            // Set gridlines to be hidden in the printed page
            pageSetup.PrintGridlines = false;

            // Set center the sheet horizontally on the page
            pageSetup.CenterHorizontally = true;

            // Set not to center the sheet vertically on the page
            pageSetup.CenterVertically = false;


            //Do not print headings
            pageSetup.PrintHeadings = false;

            // Do not print comments
            pageSetup.PrintComments = ExcelPrintLocation.PrintNoComments;

            // Do print in black and white only
            pageSetup.BlackAndWhite = true;

            // Draft quality set to false
            pageSetup.Draft = false;

            // Set first page number as 2
            pageSetup.FirstPageNumber = 2;                         

            // Save workbook as Excel document
            workbook.SaveAs("MonthlySales.xlsx");

        }
    }
}
