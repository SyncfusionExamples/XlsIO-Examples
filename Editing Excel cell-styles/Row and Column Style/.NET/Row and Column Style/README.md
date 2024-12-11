# How to apply the default style to rows and columns?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to apply the default style to rows and columns in the worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet worksheet = workbook.Worksheets[0];

	#region Row and Column Style
	//Define new styles to apply in rows and columns
	IStyle rowStyle = workbook.Styles.Add("RowStyle");
	rowStyle.Color = Syncfusion.Drawing.Color.LightGreen;
	IStyle columnStyle = workbook.Styles.Add("ColumnStyle");
	columnStyle.Color = Syncfusion.Drawing.Color.Orange;

	//Set default row style for entire row
	worksheet.SetDefaultRowStyle(1, 2, rowStyle);
	//Set default column style for entire column
	worksheet.SetDefaultColumnStyle(1, 2, columnStyle);
	#endregion

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/RowColumnStyle.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %} 