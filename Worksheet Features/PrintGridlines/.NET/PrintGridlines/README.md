# How to print the gridlines of the worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}

Step 5: Include the below code snippet in **Program.cs** to print the gridlines of the worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet sheet = workbook.Worksheets[0];

	for (int i = 1; i <= 50; i++)
	{
		for (int j = 1; j <= 50; j++)
		{
			sheet.Range[i, j].Text = sheet.Range[i, j].AddressLocal;
		}
	}

	#region PageSetup Settings
	//True to cell gridlines are printed on the page
	sheet.PageSetup.PrintGridlines = true;

	#endregion

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/PrintGridlines.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}