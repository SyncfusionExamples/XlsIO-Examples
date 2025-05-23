# How to add hyperlinks in a worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to add hyperlinks in a worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Excel2013;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet sheet = workbook.Worksheets[0];

	#region Hyperlinks
	//Creating a Hyperlink for a Website
	IHyperLink hyperlink = sheet.HyperLinks.Add(sheet.Range["C5"]);
	hyperlink.Type = ExcelHyperLinkType.Url;
	hyperlink.Address = "http://www.syncfusion.com";
	hyperlink.ScreenTip = "To know more about Syncfusion products, go through this link.";

	//Creating a Hyperlink for e-mail
	IHyperLink hyperlink1 = sheet.HyperLinks.Add(sheet.Range["C7"]);
	hyperlink1.Type = ExcelHyperLinkType.Url;
	hyperlink1.Address = "mailto:Username@syncfusion.com";
	hyperlink1.ScreenTip = "Send Mail";

	//Creating a Hyperlink for Opening Files using type as File
	IHyperLink hyperlink2 = sheet.HyperLinks.Add(sheet.Range["C9"]);
	hyperlink2.Type = ExcelHyperLinkType.File;
	hyperlink2.Address = "C:/Program files";
	hyperlink2.ScreenTip = "File path";
	hyperlink2.TextToDisplay = "Hyperlink for files using File as type";

	//Creating a Hyperlink for Opening Files using type as Unc
	IHyperLink hyperlink3 = sheet.HyperLinks.Add(sheet.Range["C11"]);
	hyperlink3.Type = ExcelHyperLinkType.Unc;
	hyperlink3.Address = "C:/Documents and Settings";
	hyperlink3.ScreenTip = "Click here for files";
	hyperlink3.TextToDisplay = "Hyperlink for files using Unc as type";

	//Creating a Hyperlink to another cell using type as Workbook
	IHyperLink hyperlink4 = sheet.HyperLinks.Add(sheet.Range["C13"]);
	hyperlink4.Type = ExcelHyperLinkType.Workbook;
	hyperlink4.Address = "Sheet1!A15";
	hyperlink4.ScreenTip = "Click here";
	hyperlink4.TextToDisplay = "Hyperlink to cell A15";
	#endregion

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/Hyperlinks.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}