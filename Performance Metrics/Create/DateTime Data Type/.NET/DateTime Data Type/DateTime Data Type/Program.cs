using System.IO;
using Syncfusion.XlsIO;

namespace DateTime_Data_Type
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                DateTime dateTime = new DateTime(1900, 1, 1);

                //Fill 100,000 rows × 50 columns with date
                for (int row = 1; row <= 100000; row++)
                {
                    for (int column = 1; column <= 50; column++)
                    {

                        //Date Time set method
                        sheet.SetValue(row, column, dateTime.ToString(), "mm/dd/yyyy");

                    }
                    dateTime = dateTime.AddDays(1);
                }


                workbook.SaveAs(@"../../../Output/Output.xlsx");
            }
        }
    }
}





