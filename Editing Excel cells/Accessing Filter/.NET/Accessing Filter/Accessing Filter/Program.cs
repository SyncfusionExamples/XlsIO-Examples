using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;

namespace Accessing_Filter
{
    class Program
    {
        static void Main(string[] args)
        {
			using (ExcelEngine excelEngine = new ExcelEngine())
			{
				IApplication application = excelEngine.Excel;
				application.DefaultVersion = ExcelVersion.Xlsx;
				FileStream inputStream = new FileStream("../../../InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
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
		}
    }
}




