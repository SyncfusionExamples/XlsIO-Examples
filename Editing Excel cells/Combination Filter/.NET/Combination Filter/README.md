# How to apply a combination filter to Excel data?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to apply a combination filter to Excel data.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
	IWorkbook workbook = application.Workbooks.Open(inputStream);
	IWorksheet worksheet = workbook.Worksheets[0];

	#region Combination Filter
	//Creating an AutoFilter in the first worksheet. Specifying the AutoFilter range. 
	worksheet.AutoFilters.FilterRange = worksheet.Range["A1:B22"];

	//Column index to which AutoFilter must be applied.
	IAutoFilter filter = worksheet.AutoFilters[0];

	//Applying Text filter to filter multiple text to get filter.
	filter.AddTextFilter(new string[] { "London", "Ireland", "Canada" });

	//Column index to which AutoFilter must be applied.
	filter = worksheet.AutoFilters[1];

	//Applying DateTime filter to filter the date based on DateTimeGroupingType.
	filter.AddDateFilter(2020, 11, 27, 0, 0, 0, DateTimeGroupingType.minute);
	#endregion

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/CombinationFilter.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
	inputStream.Dispose();
}
{% endhighlight %}
{% endtabs %} 