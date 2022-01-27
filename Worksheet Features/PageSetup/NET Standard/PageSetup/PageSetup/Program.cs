using System.IO;
using Syncfusion.XlsIO;

namespace PageSetup
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
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
                //Set Horizontal Page Breaks
                sheet.HPageBreaks.Add(sheet.Range["A5"]);
                //Set Vertical Page Breaks
                sheet.VPageBreaks.Add(sheet.Range["B5"]);

                //Set print title
                sheet.PageSetup.PrintTitleColumns = "$B:$E";
                sheet.PageSetup.PrintTitleRows = "$2:$5";

                //Set Page Orientation as Portrait or Landscape
                sheet.PageSetup.Orientation = ExcelPageOrientation.Landscape;
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("PageSetup-Settings.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("PageSetup-Settings.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
