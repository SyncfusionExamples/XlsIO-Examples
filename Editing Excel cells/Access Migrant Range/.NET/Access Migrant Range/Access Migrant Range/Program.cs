using System.IO;
using Syncfusion.XlsIO;

namespace Access_Migrant_Range
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
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Access Migrant Range
                //Getting migrant range of worksheet
                IMigrantRange migrantRange = worksheet.MigrantRange;

                //Writing data into migrant range
                for (int row = 1; row <= 5; row++)
                {
                    for (int column = 1; column <= 5; column++)
                    {
                        //Writing values
                        migrantRange.ResetRowColumn(row, column);
                        migrantRange.Text = "Test";
                    }
                }
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/MigrantRange.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("MigrantRange.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
