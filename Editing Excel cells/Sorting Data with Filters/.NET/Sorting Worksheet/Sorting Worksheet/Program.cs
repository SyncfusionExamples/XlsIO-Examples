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
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
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
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}