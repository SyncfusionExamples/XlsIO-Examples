# Bind data from a data table to a template marker using C#

The Syncfusion&reg; [.NET Excel Library](https://www.syncfusion.com/document-processing/excel-framework/net/excel-library) (XlsIO) enables you to create, read, and edit Excel documents programmatically without Microsoft Excel or interop dependencies. Using this library, you can **bind data from a data table to a template marker** using C#.

## Steps to bind data from a data table to a template marker programmatically

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
```csharp
using System;
using System.Data;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to bind data from a data table to a template marker.
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
    IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
    IWorksheet worksheet = workbook.Worksheets[0];

    //Create Template Marker Processor
    ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

    //Create an instance for Data table
    DataTable reports = new DataTable();

    //Add value to data table
    reports.Columns.Add("SalesPerson");
    reports.Columns.Add("FromDate", typeof(DateTime));
    reports.Columns.Add("ToDate", typeof(DateTime));

    reports.Rows.Add("Andy Bernard", new DateTime(2014, 09, 08), new DateTime(2014, 09, 11));
    reports.Rows.Add("Jim Halpert", new DateTime(2014, 09, 11), new DateTime(2014, 09, 15));
    reports.Rows.Add("Karen Fillippelli", new DateTime(2014, 09, 15), new DateTime(2014, 09, 20));
    reports.Rows.Add("Phyllis Lapin", new DateTime(2014, 09, 21), new DateTime(2014, 09, 25));
    reports.Rows.Add("Stanley Hudson", new DateTime(2014, 09, 26), new DateTime(2014, 09, 30));

    //Add collection to marker variable
    marker.AddVariable("Reports", reports, VariableTypeAction.DetectNumberFormat);

    //Process the markers in the template
    marker.ApplyMarkers();

    #region Save
    //Saving the workbook
    FileStream outputStream = new FileStream(Path.GetFullPath("Output/ImportDataTable.xlsx"), FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);
    #endregion

    //Dispose streams
    outputStream.Dispose();
    inputStream.Dispose();
}
```
{% endhighlight %}
{% endtabs %}

More information about binding data from a data table to a template marker can be found in this [documentation](https://help.syncfusion.com/document-processing/excel/excel-library/net/working-with-template-markers#bind-from-datatable) section.