# How to hide specific range in a worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to hide specific range in a worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet worksheet = workbook.Worksheets[0];

	IRange range = worksheet.Range["D4"];

	#region Hide single cell
	//Hiding the range ‘D4’
	worksheet.ShowRange(range, false);
	#endregion

	IRange firstRange = worksheet.Range["F6:I9"];
	IRange secondRange = worksheet.Range["C15:G20"];
	RangesCollection rangeCollection = new RangesCollection(application, worksheet);
	rangeCollection.Add(firstRange);
	rangeCollection.Add(secondRange);

	#region Hide multiple cells
	//Hiding a collection of ranges
	worksheet.ShowRange(rangeCollection, false);
	#endregion

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/HideRange.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}