using System.IO;
using Syncfusion.XlsIO;

namespace Sort_on_Font_Color
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

                #region Sort on Font Color
                //Creates the data sorter
                IDataSort sorter = workbook.CreateDataSorter();

                //Range to sort
                sorter.SortRange = worksheet.Range["A1:A11"];

                //Creates the sort field with the column index, sort based on and order by attribute
                ISortField sortField = sorter.SortFields.Add(0, SortOn.FontColor, OrderBy.OnTop);

                //Specifies the color to sort the data
                sortField.Color = Syncfusion.Drawing.Color.Red;

                //Sort based on the sort Field attribute
                sorter.Sort();

                //Creates the data sorter
                sorter = workbook.CreateDataSorter();

                //Range to sort
                sorter.SortRange = worksheet.Range["B1:B11"];

                //Creates another sort field with the column index, sort based on and order by attribute
                sortField = sorter.SortFields.Add(1, SortOn.FontColor, OrderBy.OnBottom);

                //Specifies the color to sort the data
                sortField.Color = Syncfusion.Drawing.Color.Red;

                //Sort based on the sort Field attribute
                sorter.Sort();
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/SortOnFontColor.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

            }
        }
    }
}
