# How to find all occurrences of text in a worksheet using different find options?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to find all occurrences of text in a worksheet using different find options.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
	IWorkbook workbook = application.Workbooks.Open(fileStream);
	IWorksheet worksheet = workbook.Worksheets[0];

	//Searches for the given string within the text of worksheet
	IRange[] result1 = worksheet.FindAll("Gill", ExcelFindType.Text);

	//Searches for the given string within the text of worksheet
	IRange[] result2 = worksheet.FindAll(700, ExcelFindType.Number);

	//Searches for the given string in formulas
	IRange[] result3 = worksheet.FindAll("=SUM(F10:F11)", ExcelFindType.Formula);

	//Searches for the given string in calculated value, number and text
	IRange[] result4 = worksheet.FindAll("41", ExcelFindType.Values);

	//Searches for the given string in comments
	IRange[] result5 = worksheet.FindAll("Desk", ExcelFindType.Comments);

	//Searches for the given string within the text of worksheet and case matched
	IRange[] result6 = worksheet.FindAll("Pen Set", ExcelFindType.Text, ExcelFindOptions.MatchCase);

	//Searches for the given string within the text of worksheet and the entire cell content matching to search text
	IRange[] result7 = worksheet.FindAll("5", ExcelFindType.Text, ExcelFindOptions.MatchEntireCellContent);

	//Saving the workbook as stream
	FileStream stream = new FileStream(Path.GetFullPath(@"Output/Find.xlsx"), FileMode.Create, FileAccess.ReadWrite);
	workbook.SaveAs(stream);
	stream.Dispose();
}
{% endhighlight %}
{% endtabs %} 