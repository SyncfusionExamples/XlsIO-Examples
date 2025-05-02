# How to apply sorting in the worksheet Autofilter using C#?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to apply sorting in the worksheet Autofilter.
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    FileStream inputStream = new FileStream(Path.GetFullPath("Data/Input.xlsx"), FileMode.Open, FileAccess.Read);
    IWorkbook workbook = application.Workbooks.Open(inputStream);
    IWorksheet worksheet = workbook.Worksheets[0];

    //Access sort fields from AutoFilters
    ISortFields sortFieldsCollection = worksheet.AutoFilters.DataSorter.SortFields;

    //Copy sort fields to a list
    List<ISortField> sortFields = new List<ISortField>();

    for (int i = 0; i < sortFieldsCollection.Count; i++)
    {
        sortFields.Add(sortFieldsCollection[i]);
    }

    //Remove each sort field
    foreach (ISortField sortField in sortFields)
    {
        worksheet.AutoFilters.DataSorter.SortFields.Remove(sortField);
    }

    //Now re-use the AutoFilters DataSorter
    IDataSort sorter = worksheet.AutoFilters.DataSorter;
    sorter.SortRange = worksheet.UsedRange;
    sorter.SortFields.Add(0, SortOn.Values, OrderBy.Ascending);
    sorter.Sort();

    #region Save
    FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);
    #endregion

    //Dispose streams
    outputStream.Dispose();
    inputStream.Dispose();
}
```

