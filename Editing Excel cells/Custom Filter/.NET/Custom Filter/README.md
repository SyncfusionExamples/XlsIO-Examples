# How to apply a custom filter to Excel data?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to apply a custom filter to Excel data.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
	IWorkbook workbook = application.Workbooks.Open(inputStream);
	IWorksheet worksheet = workbook.Worksheets[0];

	#region Custom Filter
	//Creating an AutoFilter in the first worksheet. Specifying the AutoFilter range
	worksheet.AutoFilters.FilterRange = worksheet.Range["A1:A11"];
	IAutoFilter filter = worksheet.AutoFilters[0];

	//Specifying first condition
	IAutoFilterCondition firstCondition = filter.FirstCondition;
	firstCondition.ConditionOperator = ExcelFilterCondition.Greater;
	firstCondition.Double = 100;

	//Specifying second condition
	IAutoFilterCondition secondCondition = filter.SecondCondition;
	secondCondition.ConditionOperator = ExcelFilterCondition.Less;
	secondCondition.Double = 200;
	#endregion

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/CustomFilter.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
	inputStream.Dispose();
}
{% endhighlight %}
{% endtabs %} 