# Complete Setup Guide - ASP.NET MVC 5 SQL Server to XLSB Export

This guide walks you through setting up and running the sample application from scratch.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Visual Studio Project Setup](#visual-studio-project-setup)
3. [Database Setup](#database-setup)
4. [NuGet Package Installation](#nuget-package-installation)
5. [Configuration](#configuration)
6. [Running the Application](#running-the-application)
7. [Testing the Export](#testing-the-export)
8. [Troubleshooting](#troubleshooting)

## Prerequisites

Before you begin, ensure you have the following installed on your Windows machine:

### Required Software

- **Visual Studio 2015 or later** (Community, Professional, or Enterprise)
  - With ASP.NET MVC 5 workload
  - With SQL Server Data Tools (SSDT) for database management
  - Download from: https://visualstudio.microsoft.com/

- **.NET Framework 4.6.2** (or higher)
  - Usually included with Visual Studio
  - Can be downloaded from: https://dotnet.microsoft.com/download/dotnet-framework

- **SQL Server 2012 or later**
  - SQL Server Express LocalDB (Recommended for development)
  - Included with Visual Studio
  - Or download SQL Server Express from: https://www.microsoft.com/en-us/sql-server/sql-server-express

- **NuGet Package Manager**
  - Included in Visual Studio
  - Accessible via: Tools → NuGet Package Manager → Package Manager Console

## Visual Studio Project Setup

### Step 1: Create a New ASP.NET MVC 5 Project

1. **Open Visual Studio**
2. **File → New → Project**
3. **Select Templates → Visual C# → Web**
4. **Choose "ASP.NET Web Application (.NET Framework)"**
5. **Name**: `ExcelDataExport`
6. **Location**: Choose your desired folder
7. **Framework**: .NET Framework 4.6.2
8. **Click OK**

### Step 2: Select MVC 5 Template

1. In the template dialog, select **"MVC"**
2. **Uncheck** "Host in the cloud" (unless you want Azure integration)
3. **Click OK**

Visual Studio will create the project structure with default MVC files.

### Step 3: Verify Project Structure

Your project should have:

```
ExcelDataExport/
├── App_Start/
│   ├── BundleConfig.cs
│   ├── FilterConfig.cs
│   └── RouteConfig.cs
├── Controllers/
│   └── HomeController.cs
├── Models/
├── Views/
│   ├── Home/
│   │   └── Index.cshtml
│   └── Web.config
├── Properties/
│   └── AssemblyInfo.cs
├── Global.asax
├── Global.asax.cs
├── packages.config
├── Web.config
└── ExcelDataExport.csproj
```

## Database Setup

### Option A: Using SQL Server LocalDB (Recommended for Development)

#### Step 1: Open SQL Server Management Studio (SSMS)

1. **Visual Studio → Tools → SQL Server → New Query**
   - Or directly open SSMS if installed separately

#### Step 2: Create the Database and Table

Run the following SQL script:

```sql
-- Create database
CREATE DATABASE SampleDB;
GO

-- Use the new database
USE SampleDB;
GO

-- Create Products table
CREATE TABLE Products (
    ProductID INT PRIMARY KEY IDENTITY(1,1),
    ProductName NVARCHAR(100) NOT NULL,
    Category NVARCHAR(50) NOT NULL,
    UnitPrice DECIMAL(10, 2) NOT NULL,
    UnitsInStock INT NOT NULL,
    DateAdded DATETIME NOT NULL DEFAULT GETDATE()
);
GO

-- Insert sample data
INSERT INTO Products (ProductName, Category, UnitPrice, UnitsInStock, DateAdded) 
VALUES 
    ('Laptop Computer', 'Electronics', 999.99, 15, GETDATE()),
    ('Wireless Mouse', 'Electronics', 29.99, 150, GETDATE()),
    ('USB-C Cable', 'Accessories', 12.99, 500, GETDATE()),
    ('Monitor 27inch', 'Electronics', 349.99, 25, GETDATE()),
    ('Keyboard Mechanical', 'Electronics', 89.99, 80, GETDATE()),
    ('Desk Lamp LED', 'Accessories', 44.99, 120, GETDATE()),
    ('Office Chair', 'Furniture', 299.99, 10, GETDATE()),
    ('Standing Desk', 'Furniture', 599.99, 5, GETDATE()),
    ('USB Hub', 'Accessories', 39.99, 200, GETDATE()),
    ('Webcam HD', 'Electronics', 79.99, 60, GETDATE());
GO

-- Verify data
SELECT * FROM Products;
GO
```

#### Step 3: Verify Connection

```sql
-- Test connection
USE SampleDB;
SELECT COUNT(*) AS 'Product Count' FROM Products;
```

You should see: **Product Count = 10**

### Option B: Using SQL Server Express Named Instance

If you're using SQL Server Express instead of LocalDB, replace the connection string.

**Connection String for SQL Server Express:**
```
Server=.\SQLEXPRESS;Database=SampleDB;Integrated Security=true;
```

### Option C: Using a Remote SQL Server

For a named instance or remote server:

```
Server=YOUR_SERVER_NAME\INSTANCE_NAME;Database=SampleDB;Integrated Security=true;
```

Or with SQL Server authentication:

```
Server=YOUR_SERVER_NAME;User ID=sa;Password=YourPassword;Database=SampleDB;
```

## NuGet Package Installation

### Step 1: Open Package Manager Console

**Tools → NuGet Package Manager → Package Manager Console**

### Step 2: Install Required Packages

Run these commands in the Package Manager Console:

```powershell
# Install Syncfusion XlsIO for ASP.NET
Install-Package Syncfusion.XlsIO.AspNet -Version 21.1.35

# Install Entity Framework (optional, for easier data access)
Install-Package EntityFramework -Version 6.4.4
```

### Step 3: Verify Installation

Check that packages are listed in `packages.config`:

```xml
<package id="Syncfusion.XlsIO.AspNet" version="21.1.35" targetFramework="net462" />
<package id="Syncfusion.Licensing" version="21.1.35" targetFramework="net462" />
<package id="Syncfusion.Compression.Base" version="21.1.35" targetFramework="net462" />
```

## Configuration

### Step 1: Update Web.config Connection String

**Location**: `ExcelDataExport/Web.config`

Find the `<connectionStrings>` section and update it:

**For LocalDB (default):**
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

### Step 2: Verify MIME Type Configuration

In `Web.config`, verify the XLSB MIME type is configured:

```xml
<system.webServer>
  <staticContent>
    <mimeMap fileExtension=".xlsb" mimeType="application/vnd.ms-excel.sheet.binary.macroEnabled" />
  </staticContent>
</system.webServer>
```

### Step 3: Build and Restore Packages

1. **Build → Clean Solution**
2. **Build → Rebuild Solution**

Visual Studio will automatically restore NuGet packages.

## Running the Application

### Step 1: Set Startup Project

1. **Right-click ExcelDataExport → Set as StartUp Project**

### Step 2: Configure IIS Express

1. **Project Properties → Web**
2. Verify **"IIS Express"** is selected
3. Set **Project URL** (e.g., `http://localhost:50001/`)

### Step 3: Start the Application

**Press F5** or **Debug → Start Debugging**

Visual Studio will:
- Compile the project
- Launch IIS Express
- Open the application in your default browser

### Step 4: Access the Application

The application should open at: **http://localhost:[port]/**

You'll see the home page with:
- Title: "SQL Server to XLSB Export"
- An "Export to XLSB" button
- Workflow and technical details

## Testing the Export

### Step 1: Click Export Button

1. On the home page, click **"Export to XLSB"** button
2. The application will:
   - Connect to SQL Server
   - Execute the Products query
   - Load data into DataTable
   - Create an Excel workbook
   - Import DataTable into worksheet
   - Export as XLSB format
   - Trigger file download

### Step 2: Verify Download

The file **`ProductsExport.xlsb`** should download to your browser's default download folder.

### Step 3: Open the File

1. **Navigate to your Downloads folder**
2. **Open ProductsExport.xlsb** with Microsoft Excel
3. Verify the data:
   - Column headers: ProductID, ProductName, Category, UnitPrice, UnitsInStock, DateAdded
   - 10 product rows with sample data
   - Formatted header row (blue background, white text, bold)

### Step 4: Verify XLSB Format

1. **Right-click ProductsExport.xlsb → Properties**
2. **Type**: Should show "Microsoft Excel Binary Worksheet"
3. **File extension**: `.xlsb`

## Troubleshooting

### Problem: "Cannot connect to database"

**Symptoms**: Red error message on the page after clicking Export

**Solutions**:

1. **Verify SQL Server is running**:
   - Open Task Manager → Services tab
   - Look for SQL Server service (starts with "SQL Server")
   - Should be "Running"

2. **Verify database exists**:
   - Open SSMS
   - Connect to your SQL Server instance
   - Expand Databases → Look for "SampleDB"
   - If missing, re-run the database creation script

3. **Verify connection string**:
   - Open `Web.config`
   - Check the `DefaultConnection` string
   - Test the connection in Visual Studio:
     - **View → Server Explorer**
     - **Add Connection → Data Source: SQL Server → Server Name: (localdb)\mssqllocaldb**
     - Should show your instance

4. **Verify Integrated Security**:
   - Ensure the connection string uses `Integrated Security=true`
   - Your Windows account must have SQL Server access rights
   - Ask your SQL Server admin to add your Windows account as a login

### Problem: "Products table not found" Error

**Solutions**:

1. Verify the table exists:
   ```sql
   USE SampleDB;
   SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Products';
   ```

2. If not found, run the table creation script again:
   ```sql
   USE SampleDB;
   CREATE TABLE Products (...)
   ```

### Problem: File doesn't download

**Solutions**:

1. **Check browser download settings**:
   - Chrome/Edge: Settings → Privacy and security → Downloads
   - Verify "Ask where to save each file" or check your Downloads folder

2. **Check MIME type**:
   - Verify `Web.config` has XLSB MIME type configured
   - Reload the page (Ctrl+F5)
   - Click Export again

3. **Clear browser cache**:
   - Press Ctrl+Shift+Delete
   - Clear cached files
   - Try again

### Problem: "Syncfusion license expiration" Warning

**Solutions**:

1. This is a trial limitation. For production use:
   - Purchase a Syncfusion license from: https://www.syncfusion.com/
   - Apply the license key in code:
     ```csharp
     Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY");
     ```

2. Trial limitations:
   - Up to 25 worksheets per workbook
   - No licensing watermark in debug mode
   - For production, license is required

### Problem: XLSB file is corrupted

**Solutions**:

1. **Verify file was created successfully**:
   - Check that no errors appear on the page
   - File size should be > 1 KB

2. **Verify Syncfusion version**:
   - Open `packages.config`
   - Confirm Syncfusion.XlsIO.AspNet version is 21.1.35 or compatible

3. **Try downloading again**:
   - Clear browser cache
   - Click Export button again
   - Try opening with a different Excel version

### Problem: Application won't start in Visual Studio

**Solutions**:

1. **Verify .NET Framework 4.6.2 is installed**:
   - **Control Panel → Programs → Programs and Features**
   - Look for ".NET Framework 4.6.2"
   - If not found, download from: https://dotnet.microsoft.com/download/dotnet-framework

2. **Rebuild project**:
   - **Build → Clean Solution**
   - **Build → Rebuild Solution**
   - **F5** to start debugging

3. **Reset IIS Express**:
   - Close Visual Studio
   - Delete `%USERPROFILE%\Documents\IISExpress\config` folder
   - Open Visual Studio → Run again

## Performance Optimization

For production deployments with large datasets:

1. **Implement pagination**: Fetch data in batches
2. **Use connection pooling**: Already enabled by default in ADO.NET
3. **Add compression**: Enable GZip in IIS
4. **Cache worksheet formatting**: If exporting frequently
5. **Consider async operations**: Use async/await for I/O

## Security Considerations

1. **Validate user input**: If adding parameters to SQL query
2. **Use parameterized queries**: Already implemented in the sample
3. **Restrict file downloads**: Add authentication/authorization
4. **Encrypt sensitive data**: Use SQL Server encryption
5. **Audit file exports**: Log who downloaded what and when

---

**Next Steps**: After successful export, you can:
- Add more data columns to the query
- Implement pagination for large datasets
- Add authentication/authorization
- Deploy to IIS or Azure
- Integrate into your existing application

For more information, see README.md and inline code comments.
