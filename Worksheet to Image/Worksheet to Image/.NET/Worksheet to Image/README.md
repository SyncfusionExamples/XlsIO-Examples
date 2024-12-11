# How to convert an Excel worksheet to an image?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to convert an Excel worksheet to an image.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
	IWorkbook workbook = application.Workbooks.Open(inputStream);
	IWorksheet sheet = workbook.Worksheets[0];

	//Initialize XlsIORenderer
	application.XlsIORenderer = new XlsIORenderer();

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/Image.png"), FileMode.Create, FileAccess.Write);
	sheet.ConvertToImage(sheet.UsedRange, outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
	inputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}