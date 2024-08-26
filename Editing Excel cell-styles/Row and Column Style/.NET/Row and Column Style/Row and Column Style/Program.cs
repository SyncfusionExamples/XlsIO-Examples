using System.IO;
using Syncfusion.XlsIO;

namespace Row_and_Column_Style
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

                #region Row and Column Style
                //Define new styles to apply in rows and columns
                IStyle rowStyle = workbook.Styles.Add("RowStyle");
                rowStyle.Color = Syncfusion.Drawing.Color.LightGreen;
                IStyle columnStyle = workbook.Styles.Add("ColumnStyle");
                columnStyle.Color = Syncfusion.Drawing.Color.Orange;

                //Set default row style for entire row
                worksheet.SetDefaultRowStyle(1, 2, rowStyle);
                //Set default column style for entire column
                worksheet.SetDefaultColumnStyle(1, 2, columnStyle);
				#endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/RowColumnStyle.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("RowColumnStyle.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
