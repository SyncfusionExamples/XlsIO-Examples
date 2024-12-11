# How to to move a row from one worksheet to another worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to move a row from one worksheet to another worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
  IApplication application = excelEngine.Excel;
  application.DefaultVersion = ExcelVersion.Xlsx;
  FileStream sourceStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
  IWorkbook workbook = application.Workbooks.Open(sourceStream);

  IWorksheet sourceWorksheet = workbook.Worksheets[0];
  IWorksheet destinationWorksheet = workbook.Worksheets[1];

  IRange sourceRow = sourceWorksheet.Range[2, 1];
  IRange destinationRow = destinationWorksheet.Range[2, 1];

  //Move the entire row to the next sheet
  sourceRow.EntireRow.MoveTo(destinationRow);

  //Saving the workbook as stream
  FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.Write);
  workbook.SaveAs(outputStream);

  //Dispose streams
  outputStream.Dispose();
  sourceStream.Dispose();
}
{% endhighlight %}
{% endtabs %}