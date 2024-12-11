# How to apply rich text formatting in the worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to apply rich text formatting in the worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet worksheet = workbook.Worksheets[0];

	//Add Text
	IRange range = worksheet.Range["A1"];
	range.Text = "RichText";
	IRichTextString richText = range.RichText;

	//Formatting first 4 characters.
	IFont redFont = workbook.CreateFont();
	redFont.Bold = true;
	redFont.Italic = true;
	redFont.RGBColor = Syncfusion.Drawing.Color.Red;
	richText.SetFont(0, 3, redFont);

	//Formatting last 4 characters.
	IFont blueFont = workbook.CreateFont();
	blueFont.Bold = true;
	blueFont.Italic = true;
	blueFont.RGBColor = Syncfusion.Drawing.Color.Blue;
	richText.SetFont(4, 7, blueFont);

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/RichText.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}