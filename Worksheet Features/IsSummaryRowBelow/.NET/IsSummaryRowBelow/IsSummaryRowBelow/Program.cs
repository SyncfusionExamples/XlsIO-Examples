using System.IO;
using Syncfusion.XlsIO;

namespace IsSummaryRowBelow
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

                for (int i = 1; i <= 50; i++)
                {
                    for (int j = 1; j <= 50; j++)
                    {
                        sheet.Range[i, j].Text = sheet.Range[i, j].AddressLocal;
                    }
                }

                #region PageSetup Settings
                //True to summary rows will appear below detail in outlines
                sheet.PageSetup.IsSummaryRowBelow = true;
                sheet.PageSetup.FitToPagesWide = 0;
                sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
                sheet.PageSetup.IsFitToPage = true;

                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("SummaryRowBelow.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("SummaryRowBelow.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
