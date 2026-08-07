# Split an Excel File into Multiple Excel Files in C#

The [.NET Excel Library](https://www.syncfusion.com/document-sdk/net-excel-library) enables you to create, read, and edit Excel documents programmatically without Microsoft Excel or interop dependencies. Using this library, you can **split an Excel file into multiple Excel files** using C#.

## Steps to Split an Excel File into Multiple Excel Files

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
```csharp
using System.IO;
using Syncfusion.XlsIO;
```
Step 5: Include the below code snippet in **Program.cs** to split an Excel file into multiple Excel files
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Open(inputData);
    IWorksheets worksheets = workbook.Worksheets;

    workbook.Version = ExcelVersion.Xlsx;

    //Loop through each Excel worksheet and save it as a new workbook
    foreach (IWorksheet worksheet in worksheets)
    {
        IWorkbook newBook = application.Workbooks.Create(0);
        newBook.Worksheets.AddCopy(worksheet);

        FileStream outputData = new FileStream(outputPath + worksheet.Name + ".xlsx", FileMode.Create, FileAccess.ReadWrite);
        newBook.SaveAs(outputData);
        outputData.Close();
    }
}
```