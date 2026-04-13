using System.IO;
using Syncfusion.XlsIO;

namespace SortWithFirstColumn
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                #region Workbook initialization
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];
                #endregion

                #region Sort On Cell Values
                //Creates the data sorter
                IDataSort sorter = workbook.CreateDataSorter();

                //This includes the first column for sorting range
                sorter.HasHeader = false;

                //Range to sort
                sorter.SortRange = worksheet.Range["D1:D26"];

                //Adds a sort field: then by values in column D in descending order
                ISortField sortField = sorter.SortFields.Add(3, SortOn.Values, OrderBy.Descending);

                //Setting the algorithm to sort the values
                sorter.Algorithm = SortingAlgorithms.QuickSort;
                sorter.Sort();
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
                #endregion
            }
        }
    }
}
