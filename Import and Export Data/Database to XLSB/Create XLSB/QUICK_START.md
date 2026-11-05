# Quick Start Guide

Get the application up and running in 5 minutes.

## Prerequisites Checklist

- [ ] Visual Studio 2015 or later installed
- [ ] .NET Framework 4.6.2 available
- [ ] SQL Server or LocalDB available
- [ ] NuGet Package Manager accessible

## 5-Minute Setup

### 1. Create Database (1 minute)

Open SQL Server Management Studio and run:

```sql
CREATE DATABASE SampleDB;
GO

USE SampleDB;
GO

CREATE TABLE Products (
    ProductID INT PRIMARY KEY IDENTITY(1,1),
    ProductName NVARCHAR(100) NOT NULL,
    Category NVARCHAR(50) NOT NULL,
    UnitPrice DECIMAL(10, 2) NOT NULL,
    UnitsInStock INT NOT NULL,
    DateAdded DATETIME NOT NULL DEFAULT GETDATE()
);
GO

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
('Webcam HD', 'Electronics', 79.99, 60, GETDATE());
GO
```

### 2. Open Project (1 minute)

1. Open Visual Studio
2. Open `ExcelDataExport.sln` (the project folder)
3. Wait for solution to load

### 3. Install NuGet Packages (2 minutes)

In **Package Manager Console** (Tools → NuGet Package Manager → Package Manager Console):

```powershell
Install-Package Syncfusion.XlsIO.AspNet -Version 21.1.35
```

Wait for installation to complete.

### 4. Update Connection String (30 seconds)

In **Web.config**, update:

```xml
<connectionStrings>
  <add name="DefaultConnection" 
       connectionString="Server=(localdb)\mssqllocaldb;Database=SampleDB;Integrated Security=true;" 
       providerName="System.Data.SqlClient" />
</connectionStrings>
```

Adjust if using SQL Server Express:
```
Server=.\SQLEXPRESS;Database=SampleDB;Integrated Security=true;
```

### 5. Run Application (30 seconds)

1. Press **F5** to start debugging
2. Browser opens to the application
3. Click **"Export to XLSB"** button
4. File downloads as `ProductsExport.xlsb`

## Verify It Works

1. ✓ Application loads without errors
2. ✓ "Export to XLSB" button is visible
3. ✓ File downloads to your Downloads folder
4. ✓ File opens in Excel showing product data

## Done! 🎉

Your ASP.NET MVC 5 application is now exporting data from SQL Server to XLSB format.

---

## Troubleshooting

### "Cannot connect to database"

**Check connection string in Web.config**

For LocalDB:
```
Server=(localdb)\mssqllocaldb;Database=SampleDB;Integrated Security=true;
```

For SQL Server Express:
```
Server=.\SQLEXPRESS;Database=SampleDB;Integrated Security=true;
```

### "Products table not found"

**Verify table exists in SQL Server:**

```sql
USE SampleDB;
SELECT * FROM Products;
```

Should return 10 rows of product data.

### File doesn't download

**Clear browser cache** (Ctrl+Shift+Delete) and try again.

---

For detailed setup instructions, see: **SETUP_GUIDE.md**

For code explanation, see: **CODE_EXPLANATION.md**

For complete documentation, see: **README.md**
