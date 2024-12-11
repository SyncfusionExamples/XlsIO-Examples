# How to access different types of filters in the worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to access different types of filters in the worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
	IWorkbook workbook = application.Workbooks.Open(inputStream);
	IWorksheet worksheet = workbook.Worksheets[0];

	#region Accessing Filter
	//selecting the filter by column index
	IAutoFilter filter = worksheet.AutoFilters[0];

	switch (filter.FilterType)
	{
		case ExcelFilterType.CombinationFilter:
			CombinationFilter filterItems = (filter.FilteredItems as CombinationFilter);
			for (int index = 0; index < filterItems.Count; index++)
			{
				if (filterItems[index].CombinationFilterType == ExcelCombinationFilterType.TextFilter)
				{
					string textValue = (filterItems[index] as TextFilter).Text;
				}
				else
				{
					DateTimeGroupingType groupType = (filterItems[index] as DateTimeFilter).GroupingType;
				}
			}
			break;

		case ExcelFilterType.DynamicFilter:
			DynamicFilter dateFilter = (filter.FilteredItems as DynamicFilter);
			DynamicFilterType dynamicFilterType = dateFilter.DateFilterType;
			break;

		case ExcelFilterType.CustomFilter:
			IAutoFilterCondition firstCondition = filter.FirstCondition;
			ExcelFilterDataType types = firstCondition.DataType;
			break;

		case ExcelFilterType.ColorFilter:
			ColorFilter colorFilter = (filter.FilteredItems as ColorFilter);
			Syncfusion.Drawing.Color color = colorFilter.Color;
			ExcelColorFilterType filterType = colorFilter.ColorFilterType;
			break;

		case ExcelFilterType.IconFilter:
			IconFilter iconFilter = (filter.FilteredItems as IconFilter);
			int iconId = iconFilter.IconId;
			ExcelIconSetType iconSetType = iconFilter.IconSetType;
			break;
	}
	#endregion

	//Dispose streams
	inputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}