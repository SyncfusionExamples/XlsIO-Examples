using System.IO;
using Syncfusion.XlsIO;

namespace Sort_On_Cell_Values
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Sort On Cell Values
                //Creates the data sorter
                IDataSort sorter = workbook.CreateDataSorter();

                //Range to sort
                sorter.SortRange = worksheet.Range["A1:B11"];

                //Adds a sort field: sort by values in column A in ascending order
                sorter.SortFields.Add(0, SortOn.Values, OrderBy.Ascending);

                //Adds a sort field: then by values in column B in descending order
                sorter.SortFields.Add(1, SortOn.Values, OrderBy.Descending);

                //Sort based on the sort Field attribute
                sorter.Sort();

                //Creates the data sorter
                sorter = workbook.CreateDataSorter();

                //Range to sort
                sorter.SortRange = worksheet.Range["C1:C11"];

                //Adds a sort field: sort by values in column C in descending order
                sorter.SortFields.Add(2, SortOn.Values, OrderBy.Descending);

                //Sort based on the sort Field attribute
                sorter.Sort();
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/SortOnValues.xlsx"));
                #endregion
            }
        }
    }
}




