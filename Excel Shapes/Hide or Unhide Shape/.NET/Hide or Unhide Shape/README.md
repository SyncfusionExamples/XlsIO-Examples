# How to hide or un-hide a shape in Excel using C#?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Shapes;
```

Step 5: Include the below code snippet in **Program.cs** to hide or un-hide a shape in Excel using C#.
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;

    FileStream inputStream = new FileStream("Data/Input.xlsx", FileMode.Open, FileAccess.Read);
    IWorkbook workbook = application.Workbooks.Open(inputStream);
    IWorksheet worksheet = workbook.Worksheets[0];

    IShapes shapes = worksheet.Shapes;
    AutoShapeImpl shape1 = shapes[0] as AutoShapeImpl;

    //Set shape1 to be hidden
    shape1.IsHidden = true;

    AutoShapeImpl shape2 = shapes[1] as AutoShapeImpl;

    //Set shape2 to be visible
    shape2.IsHidden = false;

    //Saving the workbook as stream
    FileStream outputStream = new FileStream("Output/Output.xlsx", FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);
    workbook.Close();
    excelEngine.Dispose();
}
```