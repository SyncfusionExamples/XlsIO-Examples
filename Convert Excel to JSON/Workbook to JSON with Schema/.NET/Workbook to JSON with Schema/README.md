# How to convert an Excel workbook to a JSON file with schema?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using System;
using System.IO;
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to convert an Excel workbook to a JSON file with schema.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
	IWorkbook workbook = application.Workbooks.Open(inputStream);
	IWorksheet worksheet = workbook.Worksheets[0];

	#region save as JSON
	//Saves the workbook to a JSON filestream, as schema by default
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/Excel-Workbook-To-JSON-as-schema-default.json"), FileMode.Create, FileAccess.ReadWrite);
	workbook.SaveAsJson(outputStream);

	//Saves the workbook to a JSON filestream as schema
	FileStream stream1 = new FileStream(Path.GetFullPath("Output/Excel-Workbook-To-JSON-as-schema.json"), FileMode.Create, FileAccess.ReadWrite);
	workbook.SaveAsJson(stream1, true);
	#endregion

	//Dispose streams
	outputStream.Dispose();
	stream1.Dispose();
	inputStream.Dispose();

	#region Open JSON 
	//Open default JSON

	//Open JSON with Schema
	#endregion
}
{% endhighlight %}
{% endtabs %}