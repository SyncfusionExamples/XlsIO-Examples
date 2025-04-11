# How to parse worksheets on demand in Excel using C#?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to parse worksheets on demand in Excel.
```csharp
 using (ExcelEngine excelEngine = new ExcelEngine())
 {
     IApplication application = excelEngine.Excel;
application.DefaultVersion = ExcelVersion.Xlsx;
FileStream inputStream = new FileStream("Data/Input.xlsx", FileMode.Open, FileAccess.Read);
IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic, ExcelParseOptions.ParseWorksheetsOnDemand);

// Access the first worksheet (triggers parsing)
IWorksheet worksheet = workbook.Worksheets[0];
// Process your data
string value = worksheet.Range["A1"].Text;

// Save to file system
FileStream stream = new FileStream("Output/Output.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite);
workbook.SaveAs(stream);
workbook.Close();
excelEngine.Dispose();
 }
```

