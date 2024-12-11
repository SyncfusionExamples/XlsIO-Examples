# How to group rows and columns in the worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to group rows and columns in the worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate - ToGroup.xlsx"), FileMode.Open, FileAccess.Read);
    IWorkbook workbook = application.Workbooks.Open(inputStream);
    IWorksheet worksheet = workbook.Worksheets[0];

    #region Group Rows
    //Group Rows
    worksheet.Range["A3:A7"].Group(ExcelGroupBy.ByRows, true);
    worksheet.Range["A11:A16"].Group(ExcelGroupBy.ByRows);
    #endregion

    #region Group Columns
    //Group Columns
    worksheet.Range["C1:D1"].Group(ExcelGroupBy.ByColumns, false);
    worksheet.Range["F1:G1"].Group(ExcelGroupBy.ByColumns);
    #endregion

    #region Save
    //Saving the workbook
    FileStream outputStream = new FileStream(Path.GetFullPath("Output/GroupRowsandColumns.xlsx"), FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);
    #endregion

    //Dispose streams
    outputStream.Dispose();
    inputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}