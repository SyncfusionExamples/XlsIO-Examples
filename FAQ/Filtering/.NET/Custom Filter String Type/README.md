# How to apply custom filtering to string data types using XlsIO?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to apply custom filtering to string data types using XlsIO.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Input.xlsx"), FileMode.Open, FileAccess.Read);
    IWorkbook workbook = application.Workbooks.Open(inputStream);
    IWorksheet worksheet = workbook.Worksheets[0];

    //Creating an AutoFilter 
    worksheet.AutoFilters.FilterRange = worksheet.Range["A1:A11"];
    IAutoFilter filter = worksheet.AutoFilters[0];

    //Specifying first condition
    IAutoFilterCondition firstCondition = filter.FirstCondition;
    firstCondition.ConditionOperator = ExcelFilterCondition.DoesNotContain;
    firstCondition.String = "1000.00";

    //Saving the workbook
    FileStream outputStream = new FileStream(Path.GetFullPath("../../../Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);

    //Dispose streams
    outputStream.Dispose();
    inputStream.Dispose();
}
```

