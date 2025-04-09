# How to delete hyperlinks from worksheet without affecting the cell styles?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;
using Syncfusion.XlsIO.Implementation.Collections;
```

Step 5: Include the below code snippet in **Program.cs** to delete hyperlinks from worksheet without affecting the cell styles.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
        IApplication application = excelEngine.Excel;
        application.DefaultVersion = ExcelVersion.Xlsx;
        FileStream inputStream = new FileStream("Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
        IWorkbook workbook = application.Workbooks.Open(inputStream);
        IWorksheet worksheet = workbook.Worksheets[0];

        // Remove first hyperlink without affecting cell styles
        HyperLinksCollection hyperlink = worksheet.HyperLinks as HyperLinksCollection;
        hyperlink.Remove(hyperlink[0] as HyperLinkImpl);

        //Saving the workbook as stream
        FileStream outputStream = new FileStream("Output/Output.xlsx", FileMode.Create, FileAccess.Write);
        workbook.SaveAs(outputStream);
        workbook.Close();
        excelEngine.Dispose();
}
```

