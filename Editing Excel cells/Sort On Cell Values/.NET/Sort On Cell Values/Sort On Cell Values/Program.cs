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
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Sort On Cell Values
                //Creates the data sorter
                IDataSort sorter = workbook.CreateDataSorter();

                //Range to sort
                sorter.SortRange = worksheet.UsedRange;

                //Adds a sort field: sort by values in column A in ascending order
                sorter.SortFields.Add(0, SortOn.Values, OrderBy.Ascending);

                //Adds a sort field: then by values in column B in descending order
                sorter.SortFields.Add(1, SortOn.Values, OrderBy.Descending);

                //Sort based on the sort Field attribute
                sorter.Sort();

                //Creates the data sorter
                sorter = workbook.CreateDataSorter();

                //Range to sort
                sorter.SortRange = worksheet.UsedRange;

                //Adds a sort field: sort by values in column C in descending order
                sorter.SortFields.Add(2, SortOn.Values, OrderBy.Descending);

                //Sort based on the sort Field attribute
                sorter.Sort();
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/SortOnValues.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}




