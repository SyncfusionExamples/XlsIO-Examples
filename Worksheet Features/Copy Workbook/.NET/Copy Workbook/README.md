# How to create worksheets within a workbook?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to copy an entire workbook to another workbook.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    FileStream sourceStream = new FileStream(Path.GetFullPath(@"Data/SourceWorkbookTemplate.xlsx"), FileMode.Open, FileAccess.Read);
    IWorkbook sourceWorkbook = application.Workbooks.Open(sourceStream);
    FileStream destinationStream = new FileStream(Path.GetFullPath(@"Data/DestinationWorkbookTemplate.xlsx"), FileMode.Open, FileAccess.Read);
    IWorkbook destinationWorkbook = application.Workbooks.Open(destinationStream);

    //Clone the workbook
    destinationWorkbook = sourceWorkbook.Clone();
   
    //Saving the workbook as stream
    FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
    destinationWorkbook.SaveAs(outputStream);

    //Dispose streams
    outputStream.Dispose();
    destinationStream.Dispose();
    sourceStream.Dispose();
}
{% endhighlight %}
{% endtabs %} 