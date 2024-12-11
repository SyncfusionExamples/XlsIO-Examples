# How to access a single cell or range of cells in the worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to access a single cell or range of cells in the worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet sheet = workbook.Worksheets[0];

	#region Access Cell or Range
	//Access a range by specifying cell address
	sheet.Range["A7"].Text = "Accessing a Range by specify cell address ";

	//Access a range by specifying cell row and column index
	sheet.Range[9, 1].Text = "Accessing a Range by specify cell row and column index ";

	//Access a Range by specifying using defined name
	IName name = workbook.Names.Add("Name");
	name.RefersToRange = sheet.Range["A11"];
	sheet.Range["Name"].Text = "Accessing a Range by specifying using defined name";

	//Accessing a Range of cells by specifying cells address
	sheet.Range["A13:C13"].Text = "Accessing a Range of Cells (Method 1)";

	//Accessing a Range of cells specifying cell row and column index
	sheet.Range[15, 1, 15, 3].Text = "Accessing a Range of Cells (Method 2)";
	#endregion

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/AccessCellorRange.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %} 