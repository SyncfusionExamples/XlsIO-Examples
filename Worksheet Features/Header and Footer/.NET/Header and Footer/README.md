# How to add headers and footers to the printed pages?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}

Step 5: Include the below code snippet in **Program.cs** to add headers and footers to the printed pages.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Excel2013;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet worksheet = workbook.Worksheets[0];

	//Adding values in worksheet
	worksheet.Range["A1:A600"].Text = "HelloWorld";

	//Adding text with formatting to page headers 
	worksheet.PageSetup.LeftHeader = "&KFF0000 Left Header";
	worksheet.PageSetup.CenterHeader = "&KFF0000 Center Header";
	worksheet.PageSetup.RightHeader = "&KFF0000 Right Header";

	//Adding text with formatting and image to page footers
	worksheet.PageSetup.LeftFooter = "&B &18 &K0000FF Left Footer";
	FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/image.jpg"), FileMode.Open);
	worksheet.PageSetup.CenterFooter = "&G";
	worksheet.PageSetup.CenterFooterImage = Image.FromStream(imageStream);
	worksheet.PageSetup.RightFooter = "&P &K0000FF Right Footer";

	//Saving the workbook as stream
	FileStream stream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create);
	workbook.SaveAs(stream);
	stream.Dispose();
}
{% endhighlight %}
{% endtabs %}