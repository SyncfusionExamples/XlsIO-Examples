# How to add a hyperlink to a shape in a worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to add a hyperlink to a shape in a worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet worksheet = workbook.Worksheets[0];

	#region Shape Hyperlink
	//Adding hyperlink to TextBox 
	ITextBox textBox = worksheet.TextBoxes.AddTextBox(1, 1, 100, 100);
	IHyperLink hyperlink = worksheet.HyperLinks.Add((textBox as IShape), ExcelHyperLinkType.Url, "http://www.Syncfusion.com", "click here");

	//Adding hyperlink to AutoShape
	IShape autoShape = worksheet.Shapes.AddAutoShapes(AutoShapeType.Cloud, 10, 1, 100, 100);
	hyperlink = worksheet.HyperLinks.Add(autoShape, ExcelHyperLinkType.Url, "mailto:Username@syncfusion.com", "Send Mail");

	//Adding hyperlink to picture
	IPictureShape picture = worksheet.Pictures.AddPictureAsLink(5, 5, 10, 7, Path.GetFullPath(@"Data/Image.png"));
	hyperlink = worksheet.HyperLinks.Add(picture);
	hyperlink.Type = ExcelHyperLinkType.Unc;
	hyperlink.Address = "C://Documents and Settings";
	hyperlink.ScreenTip = "Click here for files";
	#endregion

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/ShapeHyperlink.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}