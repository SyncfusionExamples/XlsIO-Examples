# Project Structure and File Reference

## Directory Layout

```
ASP.NET MVC 5 - SQL Server to XLSB Export/
│
├── README.md                          ← Start here for overview
├── QUICK_START.md                    ← 5-minute setup guide
├── SETUP_GUIDE.md                    ← Detailed step-by-step setup
├── CODE_EXPLANATION.md               ← Deep dive into the code
├── PROJECT_STRUCTURE.md              ← This file
│
├── App_Start/
│   ├── BundleConfig.cs              ← CSS/JS bundle configuration
│   ├── FilterConfig.cs              ← Global action filters
│   └── RouteConfig.cs               ← URL routing rules
│
├── Controllers/
│   └── HomeController.cs            ← Main controller with export logic
│
├── Views/
│   ├── Home/
│   │   └── Index.cshtml            ← UI with Export button
│   └── Web.config                   ← Views-specific configuration
│
├── Properties/
│   └── AssemblyInfo.cs             ← Assembly metadata
│
├── Global.asax                       ← Application startup (markup)
├── Global.asax.cs                    ← Application startup (code-behind)
├── Web.config                        ← Application configuration
├── Web.Debug.config                  ← Debug-specific configuration
├── Web.Release.config                ← Release-specific configuration
├── packages.config                   ← NuGet package dependencies
└── ExcelDataExport.csproj           ← Visual Studio project file
```

## Key Files Explained

### 1. HomeController.cs

**Location**: `Controllers/HomeController.cs`

**Purpose**: Handles HTTP requests and implements the export workflow

**Key Methods**:
- `Index()` - Displays home page
- `ExportToXlsb()` - Handles export request
- `GetProductsFromDatabase()` - Fetches data from SQL Server
- `ImportDataTableToWorksheet()` - Imports data into Excel
- `FormatWorksheet()` - Applies formatting

**Lines of Code**: ~180

**Key Concepts**:
- ADO.NET data access
- Syncfusion XlsIO API
- File download
- Error handling

### 2. Index.cshtml

**Location**: `Views/Home/Index.cshtml`

**Purpose**: User interface for the application

**Features**:
- Responsive design
- Export button
- Instructions and feature list
- Gradient background
- Error message display

**Lines of Code**: ~150

### 3. Web.config

**Location**: `Web.config` (root)

**Purpose**: Application configuration

**Key Settings**:
- Database connection string
- ASP.NET configuration
- Assembly binding redirects
- XLSB MIME type
- Entity Framework settings

**Sections**:
- `<connectionStrings>` - Database connection
- `<appSettings>` - Application settings
- `<system.web>` - ASP.NET configuration
- `<system.webServer>` - IIS configuration
- `<entityFramework>` - EF configuration
- `<runtime>` - .NET runtime settings

### 4. Global.asax & Global.asax.cs

**Location**: `Global.asax` and `Global.asax.cs`

**Purpose**: Application-level event handling

**Events**:
- `Application_Start()` - Runs when app starts
  - Registers areas
  - Registers filters
  - Registers routes
  - Registers bundles

### 5. App_Start Configuration Files

#### BundleConfig.cs
- Combines CSS files into bundles
- Combines JavaScript files into bundles
- Improves performance via minification

#### FilterConfig.cs
- Registers global action filters
- Handles errors globally
- Applies security policies

#### RouteConfig.cs
- Defines URL routing rules
- Maps URLs to controller actions
- Example: `/Home/Index` maps to `HomeController.Index()`

## File Dependencies

```
HomeController.cs
├── Web.config (connection string)
├── System.Data.SqlClient (SQL Server access)
├── System.Data (DataTable)
└── Syncfusion.XlsIO (Excel export)

Index.cshtml
└── HomeController.ExportToXlsb() (form target)

Global.asax.cs
├── App_Start/RouteConfig.cs
├── App_Start/FilterConfig.cs
└── App_Start/BundleConfig.cs
```

## Critical Configuration Points

### 1. Connection String (Web.config)

```xml
<connectionStrings>
  <add name="DefaultConnection" 
       connectionString="Server=...;Database=SampleDB;..." />
</connectionStrings>
```

**This MUST be updated** for your SQL Server instance.

### 2. XLSB MIME Type (Web.config)

```xml
<mimeMap fileExtension=".xlsb" 
         mimeType="application/vnd.ms-excel.sheet.binary.macroEnabled" />
```

Ensures browser recognizes XLSB files correctly.

### 3. Syncfusion NuGet Package (packages.config)

```xml
<package id="Syncfusion.XlsIO.AspNet" version="21.1.35" />
```

Must be installed via NuGet Package Manager.

## Workflow Diagram

```
Browser Request
    ↓
Routing (RouteConfig.cs)
    ↓
HomeController.ExportToXlsb()
    ↓
GetProductsFromDatabase()
    ├→ Read Web.config (connection string)
    ├→ Open SqlConnection
    ├→ Execute SQL query
    └→ Return DataTable
    ↓
ExcelEngine.Workbooks.Create()
    ├→ Create new workbook
    └→ Get worksheet
    ↓
ImportDataTableToWorksheet()
    ├→ Import DataTable
    └→ Create columns and rows in Excel
    ↓
FormatWorksheet()
    ├→ Format header row (bold, blue background)
    ├→ Auto-fit columns
    └→ Set minimum widths
    ↓
SaveAsStream() [XLSB Format]
    └→ Create XLSB file in memory
    ↓
File() - HTTP Response
    ├→ Set MIME type
    ├→ Set filename (ProductsExport.xlsb)
    └→ Download to browser
    ↓
Browser downloads ProductsExport.xlsb
```

## Build Configuration

### Debug vs Release

**Debug Configuration** (development):
- Compilation debug="true"
- Full debugging symbols included
- Slower execution
- Good for development

**Release Configuration** (production):
- Compilation debug="false"
- Symbols removed
- Optimizations applied
- Faster execution

Switch in Visual Studio: Build → Configuration Manager

## NuGet Dependencies

### Direct Dependencies

```
Syncfusion.XlsIO.AspNet (21.1.35)
├── Syncfusion.Core
├── Syncfusion.Compression.Base
└── Syncfusion.XlsIO.Base

Microsoft.AspNet.Mvc (5.2.7)
├── Microsoft.AspNet.WebPages
└── Microsoft.Web.Infrastructure

Microsoft.AspNet.WebPages (5.2.7)
└── System.Web.Mvc

EntityFramework (6.4.4)
```

### Transitive Dependencies

Additional packages automatically installed as dependencies.

## Assembly Binding

Redirect outdated assembly versions to correct ones:

```xml
<bindingRedirect oldVersion="0.0.0.0-12.0.0.0" newVersion="12.0.0.0" />
```

Prevents version conflicts.

## IIS Express Configuration

The application runs on IIS Express (local development server):

- **Port**: Auto-assigned (typically 50001-50100)
- **URL**: http://localhost:[port]/
- **Configuration**: `.vs\config\applicationhost.config`

## Database Table Schema

```sql
CREATE TABLE Products (
    ProductID INT PRIMARY KEY IDENTITY(1,1),
    ProductName NVARCHAR(100) NOT NULL,
    Category NVARCHAR(50) NOT NULL,
    UnitPrice DECIMAL(10, 2) NOT NULL,
    UnitsInStock INT NOT NULL,
    DateAdded DATETIME NOT NULL DEFAULT GETDATE()
);
```

**Columns**:
- `ProductID`: Unique product identifier (auto-increment)
- `ProductName`: Name of the product
- `Category`: Product category
- `UnitPrice`: Price per unit
- `UnitsInStock`: Quantity available
- `DateAdded`: When the product was added

## Deployment Notes

### Local IIS Deployment

1. Publish application from Visual Studio
2. Create Application Pool in IIS
3. Set .NET version to 4.6.2
4. Configure SQL Server connection
5. Create virtual directory

### Azure Deployment

1. Right-click project → Publish
2. Choose Azure App Service
3. Configure database connection
4. Deploy

### Security Considerations

- Use HTTPS in production
- Validate all user input
- Use parameterized SQL queries
- Implement authentication/authorization
- Protect sensitive configuration

## Performance Optimization

1. **Enable compression** in IIS
2. **Minify CSS/JS** via bundles
3. **Use pagination** for large datasets
4. **Cache data** when appropriate
5. **Optimize SQL queries** with indexes

## Logging and Diagnostics

Add logging to Web.config:

```xml
<system.diagnostics>
  <trace autoflush="true">
    <listeners>
      <add name="txtTraceListener" type="System.Diagnostics.TextWriterTraceListener" 
           initializeData="TraceLog.txt" />
    </listeners>
  </trace>
</system.diagnostics>
```

## Testing

### Unit Testing

Test individual methods:

```csharp
[TestMethod]
public void GetProductsFromDatabase_ReturnsValidDataTable()
{
    // Arrange
    var controller = new HomeController();
    
    // Act
    var result = controller.GetProductsFromDatabase();
    
    // Assert
    Assert.IsNotNull(result);
    Assert.IsTrue(result.Rows.Count > 0);
}
```

### Integration Testing

Test complete workflow with real database.

## Next Steps

1. **Review QUICK_START.md** for setup
2. **Read SETUP_GUIDE.md** for detailed instructions
3. **Study CODE_EXPLANATION.md** for code details
4. **Refer to README.md** for comprehensive documentation
5. **Customize** for your specific requirements

---

For more information, see the documentation files included in this project.
