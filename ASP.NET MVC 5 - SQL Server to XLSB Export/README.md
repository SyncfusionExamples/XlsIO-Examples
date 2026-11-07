# ASP.NET MVC 5 - SQL Server to XLSB Export

This sample demonstrates a complete customer workflow:
1. Read data from SQL Server using ADO.NET
2. Load retrieved data into a DataTable
3. Import the DataTable into an Excel worksheet using Syncfusion XlsIO
4. Export and download the workbook in XLSB format

## Prerequisites

- Visual Studio 2015 or later
- .NET Framework 4.6.2
- SQL Server 2012 or later (or SQL Server Express/LocalDB)
- NuGet Package Manager

## NuGet Packages Required

Install these packages via NuGet Package Manager Console:

```powershell
Install-Package Syncfusion.XlsIO.AspNet -Version 21.1.35
Install-Package EntityFramework -Version 6.4.4
```

**Package Details:**
- **Syncfusion.XlsIO.AspNet**: The core XlsIO library for Excel operations, including native XLSB format support
- **EntityFramework**: Optional, used for database context management

## Database Setup

### Option 1: SQL Server Express LocalDB (Recommended for Development)

1. Open SQL Server Management Studio (SSMS) or SQL Server Express
2. Create a new database named `SampleDB`
3. Run the following SQL script to create the test table and insert sample data:

```sql
USE SampleDB
GO

-- Create Products table
CREATE TABLE Products (
    ProductID INT PRIMARY KEY IDENTITY(1,1),
    ProductName NVARCHAR(100) NOT NULL,
    Category NVARCHAR(50) NOT NULL,
    UnitPrice DECIMAL(10, 2) NOT NULL,
    UnitsInStock INT NOT NULL,
    DateAdded DATETIME NOT NULL DEFAULT GETDATE()
)
GO

-- Insert sample data
INSERT INTO Products (ProductName, Category, UnitPrice, UnitsInStock, DateAdded) VALUES
('Laptop Computer', 'Electronics', 999.99, 15, GETDATE()),
('Wireless Mouse', 'Electronics', 29.99, 150, GETDATE()),
('USB-C Cable', 'Accessories', 12.99, 500, GETDATE()),
('Monitor 27inch', 'Electronics', 349.99, 25, GETDATE()),
('Keyboard Mechanical', 'Electronics', 89.99, 80, GETDATE()),
('Desk Lamp LED', 'Accessories', 44.99, 120, GETDATE()),
('Office Chair', 'Furniture', 299.99, 10, GETDATE()),
('Standing Desk', 'Furniture', 599.99, 5, GETDATE()),
('USB Hub', 'Accessories', 39.99, 200, GETDATE()),
('Webcam HD', 'Electronics', 79.99, 60, GETDATE())
GO

-- Verify data
SELECT * FROM Products
GO
```

### Option 2: Connection String Configuration

Update the connection string in `Web.config`:

```xml
<connectionStrings>
  <add name="DefaultConnection" 
       connectionString="Server=(localdb)\mssqllocaldb;Database=SampleDB;Integrated Security=true;" 
       providerName="System.Data.SqlClient" />
</connectionStrings>
```

**For SQL Server Express:**
```xml
<add name="DefaultConnection" 
     connectionString="Server=.\SQLEXPRESS;Database=SampleDB;Integrated Security=true;" 
     providerName="System.Data.SqlClient" />
```

**For SQL Server (named instance):**
```xml
<add name="DefaultConnection" 
     connectionString="Server=YOUR_SERVER_NAME;Database=SampleDB;Integrated Security=true;" 
     providerName="System.Data.SqlClient" />
```

## Project Structure

```
ExcelDataExport/
├── Controllers/
│   └── HomeController.cs
├── Models/
│   └── (Entity models if using EF)
├── Views/
│   └── Home/
│       └── Index.cshtml
├── Web.config
├── packages.config
└── ExcelDataExport.csproj
```

## How It Works

### 1. HomeController.cs

The controller handles the export request:
- Connects to SQL Server using ADO.NET
- Executes a query to fetch product data
- Loads data into a DataTable
- Creates an Excel workbook with XlsIO
- Imports the DataTable into the worksheet
- Saves the workbook in XLSB format
- Returns the file to the browser for download

### 2. Index.cshtml

The view provides a simple UI with:
- An "Export to XLSB" button
- Display of any messages or errors
- Bootstrap styling for modern appearance

## Building and Running

1. **Open the solution** in Visual Studio
2. **Restore NuGet packages**: Right-click the solution → "Restore NuGet Packages"
3. **Verify database connection**: Update `Web.config` if needed
4. **Run the application**: Press F5 or Ctrl+Shift+B
5. **Navigate to the Home page** and click "Export to XLSB"
6. The workbook will download as `ProductsExport.xlsb`

## Key Features

- **Native XLSB Export**: Uses Syncfusion XlsIO's native XLSB save API
- **ADO.NET Data Access**: Direct SQL Server connection without ORM overhead
- **DataTable Integration**: Seamless import from DataTable to Excel worksheet
- **File Download**: Proper HTTP response headers for browser download
- **Error Handling**: Try-catch blocks for database and file operations

## Troubleshooting

### "Cannot connect to database" Error
- Verify SQL Server is running
- Check connection string in `Web.config`
- Ensure database and table exist
- Confirm SQL Server authentication (Integrated or SQL Server login)

### "XlsIO license" Warning
- XlsIO trial has a 25-worksheet limitation per workbook
- For production, purchase a license from Syncfusion

### File Download Issues
- Verify MIME type registration for `.xlsb` in IIS
- Check browser download settings
- Clear browser cache and retry

## XLSB Format Advantages

- **Binary format**: More efficient than XML-based XLSX
- **Faster I/O**: Better performance with large datasets
- **Smaller file size**: Approximately 30-50% smaller than XLSX
- **Full Excel compatibility**: Supported by Excel 2007 and later

## Additional Resources

- [Syncfusion XlsIO Documentation](https://help.syncfusion.com/file-formats/xlsio/overview)
- [ASP.NET MVC 5 Documentation](https://docs.microsoft.com/en-us/aspnet/mvc/overview/getting-started)
- [ADO.NET Documentation](https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/)

---

**Author**: Bold Create  
**Framework**: ASP.NET MVC 5  
**.NET Target**: Framework 4.6.2  
**Created**: 2024
