# How to apply conditional formatting for above or below average values?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using System.IO;
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to apply conditional formatting for values below average.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
	IWorkbook workbook = application.Workbooks.Open(inputStream);
	IWorksheet worksheet = workbook.Worksheets[0];

	//Applying conditional formatting to "M6:M35"
	IConditionalFormats formats = worksheet.Range["M6:M35"].ConditionalFormats;
	IConditionalFormat format = formats.AddCondition();

	//Applying above or below average rule in the conditional formatting
	format.FormatType = ExcelCFType.AboveBelowAverage;
	IAboveBelowAverage aboveBelowAverage = format.AboveBelowAverage;

	//Set AverageType as Below for AboveBelowAverage rule.
	aboveBelowAverage.AverageType = ExcelCFAverageType.Below;

	//Set color for Conditional Formattting.
	format.FontColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 255, 255);
	format.BackColorRGB = Syncfusion.Drawing.Color.FromArgb(166, 59, 38);

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/AboveAndBelowAverage.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
	inputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}