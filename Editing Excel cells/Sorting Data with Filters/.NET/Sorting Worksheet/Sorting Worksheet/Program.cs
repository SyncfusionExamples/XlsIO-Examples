using System;
using System.IO;
using Syncfusion.XlsIO;

namespace Sorting_Worksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Access sort fields from AutoFilters
                ISortFields sortFieldsCollection = worksheet.AutoFilters.DataSorter.SortFields;

                //Copy sort fields to a list
                List<ISortField> sortFields = new List<ISortField>();

                for (int i = 0; i < sortFieldsCollection.Count; i++)
                {
                    sortFields.Add(sortFieldsCollection[i]);
                }

                //Remove each sort field
                foreach (ISortField sortField in sortFields)
                {
                    worksheet.AutoFilters.DataSorter.SortFields.Remove(sortField);
                }

                //Now re-use the AutoFilters DataSorter
                IDataSort sorter = worksheet.AutoFilters.DataSorter;
                sorter.SortRange = worksheet.UsedRange;
                sorter.SortFields.Add(0, SortOn.Values, OrderBy.Ascending);
                sorter.Sort();

                #region Save
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
                #endregion
            }
        }
    }
}