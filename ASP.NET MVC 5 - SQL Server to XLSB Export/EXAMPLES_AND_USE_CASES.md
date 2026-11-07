# Examples and Use Cases

This document provides practical examples and scenarios for extending the sample application.

## Use Case 1: Add Filtering to Export

Allow users to export only products in a specific category.

### Modify HomeController.cs

```csharp
public ActionResult ExportToXlsb(string category = null)
{
    try
    {
        // Pass category to filter
        DataTable productsTable = GetProductsFromDatabase(category);

        if (productsTable == null || productsTable.Rows.Count == 0)
        {
            ViewBag.Message = "No data found for the selected category.";
            return View("Index");
        }

        // ... rest of export logic
    }
    catch (Exception ex)
    {
        ViewBag.Message = $"Export error: {ex.Message}";
        return View("Index");
    }
}

private DataTable GetProductsFromDatabase(string category = null)
{
    DataTable dataTable = new DataTable("Products");
    string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

    using (SqlConnection connection = new SqlConnection(connectionString))
    {
        try
        {
            connection.Open();

            // Build query with optional category filter
            string query = @"
                SELECT 
                    ProductID,
                    ProductName,
                    Category,
                    UnitPrice,
                    UnitsInStock,
                    DateAdded
                FROM Products";

            if (!string.IsNullOrEmpty(category))
            {
                query += " WHERE Category = @category";
            }

            query += " ORDER BY ProductID";

            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.CommandTimeout = 30;

                // Add category parameter if provided
                if (!string.IsNullOrEmpty(category))
                {
                    command.Parameters.AddWithValue("@category", category);
                }

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
```

### Modify Index.cshtml View

```html
@using (Html.BeginForm("ExportToXlsb", "Home", FormMethod.Post))
{
    <div class="form-group">
        <label for="category">Select Category (optional):</label>
        <select name="category" id="category" class="form-control">
            <option value="">All Categories</option>
            <option value="Electronics">Electronics</option>
            <option value="Accessories">Accessories</option>
            <option value="Furniture">Furniture</option>
        </select>
    </div>

    <button type="submit" class="btn btn-primary">
        <span>📊</span>
        <span>Export to XLSB</span>
    </button>
}
```

**Result**: Users can select a category to export only matching products.

---

## Use Case 2: Multiple Worksheets

Export data to multiple worksheets in a single workbook.

### Code Example

```csharp
public ActionResult ExportMultipleCategories()
{
    try
    {
        using (IExcelEngine excelEngine = new ExcelEngine())
        {
            IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);
            workbook.Worksheets.RemoveAt(0);  // Remove default sheet

            // Get all unique categories
            List<string> categories = GetCategories();

            foreach (string category in categories)
            {
                DataTable categoryData = GetProductsFromDatabase(category);

                if (categoryData.Rows.Count > 0)
                {
                    IWorksheet worksheet = workbook.Worksheets.Create(category);
                    worksheet.ImportDataTable(categoryData, true, 1, 1);
                    FormatWorksheet(worksheet);
                }
            }

            // Also create a summary sheet
            IWorksheet summarySheet = workbook.Worksheets.Create("Summary");
            CreateSummarySheet(summarySheet);

            byte[] excelBytes = workbook.SaveAsStream().ToArray();

            return File(excelBytes, 
                        "application/vnd.ms-excel.sheet.binary.macroEnabled", 
                        "ProductsExportByCategory.xlsb");
        }
    }
    catch (Exception ex)
    {
        ViewBag.Message = $"Export error: {ex.Message}";
        return View("Index");
    }
}

private List<string> GetCategories()
{
    List<string> categories = new List<string>();
    string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

    using (SqlConnection connection = new SqlConnection(connectionString))
    {
        connection.Open();
        string query = "SELECT DISTINCT Category FROM Products ORDER BY Category";

        using (SqlCommand command = new SqlCommand(query, connection))
        {
            using (SqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    categories.Add(reader["Category"].ToString());
                }
            }
        }
    }

    return categories;
}

private void CreateSummarySheet(IWorksheet worksheet)
{
    // Create header
    worksheet.Range[1, 1].Text = "Category Summary";
    worksheet.Range[1, 1].CellStyle.Font.Bold = true;
    worksheet.Range[1, 1].CellStyle.Font.Size = 14;

    // Create column headers
    worksheet.Range[3, 1].Text = "Category";
    worksheet.Range[3, 2].Text = "Product Count";
    worksheet.Range[3, 3].Text = "Total Stock";
    worksheet.Range[3, 4].Text = "Average Price";

    // Format header row
    IRange headerRange = worksheet.Range[3, 1, 3, 4];
    headerRange.CellStyle.Font.Bold = true;
    headerRange.CellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);

    // Add summary data from database
    int row = 4;
    string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

    using (SqlConnection connection = new SqlConnection(connectionString))
    {
        connection.Open();
        string query = @"
            SELECT 
                Category,
                COUNT(*) as ProductCount,
                SUM(UnitsInStock) as TotalStock,
                AVG(UnitPrice) as AvgPrice
            FROM Products
            GROUP BY Category
            ORDER BY Category";

        using (SqlCommand command = new SqlCommand(query, connection))
        {
            using (SqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    worksheet.Range[row, 1].Text = reader["Category"].ToString();
                    worksheet.Range[row, 2].Number = (int)reader["ProductCount"];
                    worksheet.Range[row, 3].Number = (int)reader["TotalStock"];
                    worksheet.Range[row, 4].Number = Math.Round((decimal)reader["AvgPrice"], 2);
                    row++;
                }
            }
        }
    }

    // Auto-fit columns
    for (int i = 1; i <= 4; i++)
    {
        worksheet.AutofitColumn(i);
    }
}
```

**Result**: One XLSB file with separate worksheets for each category plus a summary sheet.

---

## Use Case 3: Add Calculated Columns

Export data with calculated columns (Total Revenue, Stock Value, etc.).

### Code Example

```csharp
public ActionResult ExportWithCalculations()
{
    try
    {
        DataTable productsTable = GetProductsFromDatabase();

        // Add calculated columns to DataTable
        productsTable.Columns.Add("StockValue", typeof(decimal));
        productsTable.Columns.Add("Status", typeof(string));

        foreach (DataRow row in productsTable.Rows)
        {
            decimal unitPrice = (decimal)row["UnitPrice"];
            int unitsInStock = (int)row["UnitsInStock"];

            row["StockValue"] = unitPrice * unitsInStock;
            row["Status"] = unitsInStock < 20 ? "Low Stock" : "Good Stock";
        }

        using (IExcelEngine excelEngine = new ExcelEngine())
        {
            IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Products with Calculations";

            worksheet.ImportDataTable(productsTable, true, 1, 1);
            FormatWorksheet(worksheet);

            // Format calculated columns
            IRange stockValueColumn = worksheet.Range[2, 7, worksheet.UsedRange.LastRow, 7];
            stockValueColumn.NumberFormat = "$#,##0.00";

            byte[] excelBytes = workbook.SaveAsStream().ToArray();

            return File(excelBytes, 
                        "application/vnd.ms-excel.sheet.binary.macroEnabled", 
                        "ProductsWithCalculations.xlsb");
        }
    }
    catch (Exception ex)
    {
        ViewBag.Message = $"Export error: {ex.Message}";
        return View("Index");
    }
}
```

**Result**: Excel file includes calculated columns like Stock Value and Status.

---

## Use Case 4: Add Charts and Pivot Tables

Export data with visual analysis.

### Code Example

```csharp
public ActionResult ExportWithChart()
{
    try
    {
        DataTable productsTable = GetProductsFromDatabase();

        using (IExcelEngine excelEngine = new ExcelEngine())
        {
            IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);
            IWorksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Import data
            dataSheet.ImportDataTable(productsTable, true, 1, 1);
            FormatWorksheet(dataSheet);

            // Create chart sheet
            IWorksheet chartSheet = workbook.Worksheets.Create("Chart");

            // Create pie chart by category
            IChartShape chart = chartSheet.Charts.Add();
            chart.ChartType = ExcelChartType.Pie;
            chart.DataRange = dataSheet.Range["A1:B11"];  // Category and Price
            chart.ChartTitle = "Products by Category";
            chart.TopRow = 2;
            chart.LeftColumn = 2;
            chart.BottomRow = 12;
            chart.RightColumn = 8;

            byte[] excelBytes = workbook.SaveAsStream().ToArray();

            return File(excelBytes, 
                        "application/vnd.ms-excel.sheet.binary.macroEnabled", 
                        "ProductsWithChart.xlsb");
        }
    }
    catch (Exception ex)
    {
        ViewBag.Message = $"Export error: {ex.Message}";
        return View("Index");
    }
}
```

**Result**: Excel file includes both data and visual charts.

---

## Use Case 5: Email Export Instead of Download

Send the exported file via email instead of download.

### Add NuGet Package

```powershell
Install-Package SendGrid -Version 9.28.1
```

### Code Example

```csharp
public async Task<ActionResult> ExportAndEmail(string emailAddress)
{
    try
    {
        DataTable productsTable = GetProductsFromDatabase();

        using (IExcelEngine excelEngine = new ExcelEngine())
        {
            IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Products";

            worksheet.ImportDataTable(productsTable, true, 1, 1);
            FormatWorksheet(worksheet);

            byte[] excelBytes = workbook.SaveAsStream().ToArray();

            // Send email
            await SendExcelViaEmail(emailAddress, excelBytes);

            ViewBag.Message = $"Export sent to {emailAddress}";
            return View("Index");
        }
    }
    catch (Exception ex)
    {
        ViewBag.Message = $"Error: {ex.Message}";
        return View("Index");
    }
}

private async Task SendExcelViaEmail(string recipientEmail, byte[] fileBytes)
{
    string sendGridApiKey = System.Configuration.ConfigurationManager.AppSettings["SendGridApiKey"];
    var client = new SendGridClient(sendGridApiKey);

    var from = new EmailAddress("noreply@example.com", "Excel Export Service");
    var subject = "Your Product Export";
    var to = new EmailAddress(recipientEmail);
    var plainTextContent = "Please find your product export attached.";
    var htmlContent = "<p>Please find your product export attached.</p>";

    var msg = new SendGridMessage()
    {
        From = from,
        Subject = subject,
        PlainTextContent = plainTextContent,
        HtmlContent = htmlContent
    };

    msg.AddTo(to);

    var attachment = Convert.ToBase64String(fileBytes);
    msg.AddAttachment("ProductsExport.xlsb", attachment, "application/vnd.ms-excel.sheet.binary.macroEnabled");

    await client.SendEmailAsync(msg);
}
```

**Result**: Exported file is emailed to user instead of downloading.

---

## Use Case 6: Scheduled/Batch Export

Generate exports on a schedule (daily, weekly, monthly).

### Using Hangfire for Background Jobs

Install NuGet:

```powershell
Install-Package Hangfire -Version 1.7.28
```

### Code Example

```csharp
// In App_Start\FilterConfig.cs or Global.asax.cs
using Hangfire;

public class FilterConfig
{
    public static void RegisterGlobalFilters(GlobalFilterCollection filters)
    {
        filters.Add(new HandleErrorAttribute());

        // Configure Hangfire
        GlobalConfiguration.Configuration
            .UseSqlServerStorage("DefaultConnection");

        // Schedule daily export
        RecurringJob.AddOrUpdate("DailyProductExport", 
            () => ExportAndSaveDaily(), 
            Cron.Daily(2, 0));  // 2 AM daily
    }
}

// In HomeController
public void ExportAndSaveDaily()
{
    try
    {
        DataTable productsTable = GetProductsFromDatabase();

        using (IExcelEngine excelEngine = new ExcelEngine())
        {
            IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Products";

            worksheet.ImportDataTable(productsTable, true, 1, 1);
            FormatWorksheet(worksheet);

            // Save to server
            string fileName = $"Products_{DateTime.Now:yyyyMMdd_HHmmss}.xlsb";
            string filePath = Server.MapPath($"~/Exports/{fileName}");

            workbook.SaveAsStream().WriteTo(new FileStream(filePath, FileMode.Create));
        }
    }
    catch (Exception ex)
    {
        // Log error
    }
}
```

**Result**: Automatic exports run on a schedule.

---

## Use Case 7: Large Dataset Export with Pagination

Handle large datasets efficiently by splitting into multiple files or worksheets.

### Code Example

```csharp
public ActionResult ExportLargeDataset(int pageSize = 10000)
{
    try
    {
        using (IExcelEngine excelEngine = new ExcelEngine())
        {
            IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);
            workbook.Worksheets.RemoveAt(0);

            int pageNumber = 1;
            int worksheetCount = 0;
            const int maxWorksheetsPerFile = 10;

            while (worksheetCount < maxWorksheetsPerFile)
            {
                DataTable pageData = GetProductsPagedFromDatabase(pageNumber, pageSize);

                if (pageData.Rows.Count == 0)
                    break;

                IWorksheet worksheet = workbook.Worksheets.Create($"Data_Page_{pageNumber}");
                worksheet.ImportDataTable(pageData, true, 1, 1);
                FormatWorksheet(worksheet);

                pageNumber++;
                worksheetCount++;
            }

            byte[] excelBytes = workbook.SaveAsStream().ToArray();

            return File(excelBytes, 
                        "application/vnd.ms-excel.sheet.binary.macroEnabled", 
                        "ProductsExportLarge.xlsb");
        }
    }
    catch (Exception ex)
    {
        ViewBag.Message = $"Export error: {ex.Message}";
        return View("Index");
    }
}

private DataTable GetProductsPagedFromDatabase(int pageNumber, int pageSize)
{
    DataTable dataTable = new DataTable("Products");
    int offset = (pageNumber - 1) * pageSize;

    string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

    using (SqlConnection connection = new SqlConnection(connectionString))
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
            ORDER BY ProductID
            OFFSET @offset ROWS
            FETCH NEXT @pageSize ROWS ONLY";

        using (SqlCommand command = new SqlCommand(query, connection))
        {
            command.Parameters.AddWithValue("@offset", offset);
            command.Parameters.AddWithValue("@pageSize", pageSize);
            command.CommandTimeout = 60;

            using (SqlDataAdapter adapter = new SqlDataAdapter(command))
            {
                adapter.Fill(dataTable);
            }
        }
    }

    return dataTable;
}
```

**Result**: Large datasets split into multiple worksheets.

---

## Use Case 8: Template-Based Export

Use a pre-formatted template and fill it with data.

### SQL Query

```sql
-- Create template records
INSERT INTO ProductTemplates (TemplateName, Header, Footer, HeaderColor)
VALUES ('Standard', 'Product Inventory Report', 'Confidential', 'Blue')
```

### Code Example

```csharp
public ActionResult ExportWithTemplate(string templateName)
{
    try
    {
        var template = GetTemplate(templateName);
        DataTable productsTable = GetProductsFromDatabase();

        using (IExcelEngine excelEngine = new ExcelEngine())
        {
            IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];

            // Add header
            worksheet.Range[1, 1].Text = template.Header;
            worksheet.Range[1, 1].CellStyle.Font.Bold = true;
            worksheet.Range[1, 1].CellStyle.Font.Size = 16;

            // Add data
            worksheet.ImportDataTable(productsTable, true, 3, 1);

            // Add footer
            int lastRow = 3 + productsTable.Rows.Count + 2;
            worksheet.Range[lastRow, 1].Text = template.Footer;
            worksheet.Range[lastRow, 1].CellStyle.Font.Italic = true;

            byte[] excelBytes = workbook.SaveAsStream().ToArray();

            return File(excelBytes, 
                        "application/vnd.ms-excel.sheet.binary.macroEnabled", 
                        "ProductsExport_Template.xlsb");
        }
    }
    catch (Exception ex)
    {
        ViewBag.Message = $"Export error: {ex.Message}";
        return View("Index");
    }
}

private Template GetTemplate(string templateName)
{
    // Retrieve template from database or configuration
    return new Template
    {
        Header = "Product Inventory Report",
        Footer = "Confidential - For Internal Use Only"
    };
}
```

**Result**: Export uses pre-defined templates for consistent formatting.

---

## Use Case 9: Conditional Formatting

Apply colors and formatting based on data values.

### Code Example

```csharp
private void ApplyConditionalFormatting(IWorksheet worksheet)
{
    for (int row = 2; row <= worksheet.UsedRange.LastRow; row++)
    {
        int unitsInStock = int.Parse(worksheet.Range[row, 5].Value.ToString());
        decimal unitPrice = decimal.Parse(worksheet.Range[row, 4].Value.ToString());

        IRange cell = worksheet.Range[row, 5];

        if (unitsInStock < 20)
        {
            // Low stock - red background
            cell.CellStyle.BackColor = System.Drawing.Color.Red;
            cell.CellStyle.Font.Color = ExcelKnownColors.White;
        }
        else if (unitsInStock < 50)
        {
            // Medium stock - yellow background
            cell.CellStyle.BackColor = System.Drawing.Color.Yellow;
        }
        else
        {
            // Good stock - green background
            cell.CellStyle.BackColor = System.Drawing.Color.LightGreen;
        }

        // Apply price formatting
        worksheet.Range[row, 4].NumberFormat = "$#,##0.00";
    }
}
```

**Result**: Cells are colored based on stock levels and prices are formatted as currency.

---

## Use Case 10: Export with Subtotals and Grouping

Export data with grouped rows and subtotals.

### Code Example

```csharp
private void AddSubtotals(IWorksheet worksheet, DataTable dataTable)
{
    int row = 2;
    string currentCategory = "";
    int categoryStartRow = 2;
    decimal categoryTotal = 0;

    foreach (DataRow dataRow in dataTable.Rows)
    {
        string category = dataRow["Category"].ToString();
        decimal unitPrice = (decimal)dataRow["UnitPrice"];
        int units = (int)dataRow["UnitsInStock"];

        if (currentCategory != "" && currentCategory != category)
        {
            // Insert subtotal row
            worksheet.InsertRow(row);
            worksheet.Range[row, 1].Text = $"Subtotal - {currentCategory}";
            worksheet.Range[row, 4].Number = categoryTotal;
            worksheet.Range[row, 1, row, 4].CellStyle.Font.Bold = true;
            worksheet.Range[row, 1, row, 4].CellStyle.BackColor = System.Drawing.Color.LightGray;

            categoryTotal = 0;
            categoryStartRow = row + 1;
            row++;
        }

        currentCategory = category;
        categoryTotal += unitPrice * units;
        row++;
    }

    // Add final subtotal
    worksheet.InsertRow(row);
    worksheet.Range[row, 1].Text = $"Subtotal - {currentCategory}";
    worksheet.Range[row, 4].Number = categoryTotal;
}
```

**Result**: Export includes subtotal rows with grouped data.

---

## Performance Comparison

### Small Dataset (< 1,000 rows)
- **Execution Time**: ~200ms
- **File Size**: 10-50 KB
- **Recommended**: Direct export

### Medium Dataset (1,000-50,000 rows)
- **Execution Time**: 500ms-2s
- **File Size**: 50-500 KB
- **Recommended**: Single file with pagination

### Large Dataset (50,000+ rows)
- **Execution Time**: 2s+
- **File Size**: 500 KB-10 MB
- **Recommended**: Multiple worksheets or files

---

## Best Practices

1. **Always validate user input**
2. **Use parameterized queries**
3. **Implement pagination for large datasets**
4. **Add proper error handling**
5. **Log export activities**
6. **Limit file access** to authenticated users
7. **Clean up temporary files**
8. **Monitor performance**
9. **Test with real data volumes**
10. **Document customizations**

---

For more information, see CODE_EXPLANATION.md and SETUP_GUIDE.md.
