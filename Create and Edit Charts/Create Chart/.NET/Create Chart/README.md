# How to create a chart in the worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using System.IO;
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to create a chart in the worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
	IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
	IWorksheet sheet = workbook.Worksheets[0];

	//Create a Chart
	IChartShape chart = sheet.Charts.Add();

	//Set Chart Type
	chart.ChartType = ExcelChartType.Column_Clustered;

	//Set data range in the worksheet
	chart.DataRange = sheet.Range["A1:C6"];
	chart.IsSeriesInRows = false;

	//Set Datalabels
	IChartSerie serie1 = chart.Series[0];
	IChartSerie serie2 = chart.Series[1];

	serie1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
	serie2.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
	serie1.DataPoints.DefaultDataPoint.DataLabels.Position = ExcelDataLabelPosition.Outside;
	serie2.DataPoints.DefaultDataPoint.DataLabels.Position = ExcelDataLabelPosition.Outside;

	//Set Legend
	chart.HasLegend = true;
	chart.Legend.Position = ExcelLegendPosition.Bottom;

	//Positioning the chart in the worksheet
	chart.TopRow = 8;
	chart.LeftColumn = 1;
	chart.BottomRow = 23;
	chart.RightColumn = 8;

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/Chart.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
	inputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}