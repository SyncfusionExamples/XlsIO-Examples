# How to access a relative range of cells in the worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to access a relative range of cells in the worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;

	//Setting range index mode to relative
	application.RangeIndexerMode = ExcelRangeIndexerMode.Relative;

	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet sheet = workbook.Worksheets[0];

	#region Access Reative Range
	//Creating a range by specifying cells address
	IRange range1 = sheet.Range["B3:D5"];

	//Accessing a range relatively to the existing range by specifying cell row and column index
	range1[2, 2].Text = "Returns C4 cell";
	range1[0, 0].Text = "Returns A2 cell";

	//Creating a range of cells specifying cell row and column index
	IRange range2 = sheet.Range[5, 1, 10, 3];

	//Accessing a range relatively to the existing range of cells by specifying cell row and column index
	range2[2, 2, 3, 3].Text = "Returns range of cells B6 to C7";
	#endregion

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/AccessRelativeRange.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %} 