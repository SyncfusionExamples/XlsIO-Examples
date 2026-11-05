# Access (.mdb) -> DataTable -> Excel (XLSB) in ASP.NET MVC 5

Reads the `Products` table from an Access `.mdb` file into a `DataTable`, imports it into an Excel
worksheet with Syncfusion XlsIO, and downloads it as a real `.xlsb` file.

## Setup
1. Open `ExcelDataExport.sln` in Visual Studio and restore NuGet packages.
2. Requires the Microsoft Access Database Engine (ACE OLEDB 12.0). It is installed with Microsoft Office;
   the bitness must match the web server process (IIS Express is 64-bit by default).
3. Add your Syncfusion license key in `Global.asax.cs` (`RegisterLicense`), otherwise the file carries a trial notice.
4. Press F5 and click **Export to XLSB**. `App_Data\SampleDB.mdb` is created and filled with 10 rows on first use.

## Key code (HomeController.ExportToXlsb)
```csharp
DataTable table = AccessDatabase.GetProducts();
worksheet.ImportDataTable(table, true, 1, 1);
workbook.SaveAs(stream, ExcelSaveType.SaveAsXLSB);
```
Note: XlsIO supports only basic content in XLSB (numbers, strings, dates, booleans, formulas, basic cell styles).
