# How to replace all occurrences of text with different data?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to replace all occurrences of text with different data.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
	IWorkbook workbook = application.Workbooks.Open(fileStream);
	IWorksheet worksheet = workbook.Worksheets[0];

	//Replaces the given string with another string
	worksheet.Replace("Wilson", "William");

	//Replaces the given string with another string on match case
	worksheet.Replace("4.99", "4.90", ExcelFindOptions.MatchCase);

	//Replaces the given string with another string matching entire cell content to the search word
	worksheet.Replace("Pen Set", "Pen", ExcelFindOptions.MatchEntireCellContent);

	//Replaces the given string with DateTime value
	worksheet.Replace("DateValue",DateTime.Now);

	//Replaces the given string with Array
	worksheet.Replace("Central", new string[] { "Central", "East" }, true);

	//Saving the workbook as stream
	FileStream stream = new FileStream(Path.GetFullPath("Output/Replace.xlsx"), FileMode.Create, FileAccess.ReadWrite);
	workbook.Version = ExcelVersion.Xlsx;
	workbook.SaveAs(stream);
	stream.Dispose();
}
{% endhighlight %}
{% endtabs %} 