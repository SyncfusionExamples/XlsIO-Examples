# Merge Multiple Excel Files into One Excel File in C#

The [.NET Excel Library](https://www.syncfusion.com/document-sdk/net-excel-library) enables you to create, read, and edit Excel documents programmatically without Microsoft Excel or interop dependencies. Using this library, you can **merge multiple Excel files into one Excel file** using C#.

## Steps to Merge Multiple Excel Files into One Excel File

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
```csharp
using System.IO;
using Syncfusion.XlsIO;
```
Step 5: Include the below code snippet in **Program.cs** to merge multiple Excel files into one Excel file
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(0);

    //Loop through each Excel document and add the worksheets to the new workbook
    foreach (Stream stream in streams)
    {
        stream.Position = 0;
        IWorkbook tempWorkbook = application.Workbooks.Open(stream);
        workbook.Worksheets.AddCopy(tempWorkbook.Worksheets);
        tempWorkbook.Close();
    }

    //Save the workbook to a memory stream
    MemoryStream memoryStream = new MemoryStream();
    workbook.Version = ExcelVersion.Xlsx;
    workbook.SaveAs(memoryStream);

    return memoryStream;
}
```