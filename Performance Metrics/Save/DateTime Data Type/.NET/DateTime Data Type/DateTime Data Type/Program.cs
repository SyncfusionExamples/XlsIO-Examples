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

                //Fill 150 rows × 10,000 columns with date data
                for (int row = 1; row <= 150; row++)
                {
                    for (int col = 1; col <= 10000; col++)
                    {
                        sheet[row, col].DateTime = new DateTime(2025, 1, 1).AddDays(col);
                    }
                }

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("../../../Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}





