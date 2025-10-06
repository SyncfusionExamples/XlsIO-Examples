using System.IO;
using Syncfusion.XlsIO;

namespace PrintArea
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
                //Sets the range to be printed
                sheet.PageSetup.PrintArea = "A1:M20";

                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/PrintArea.xlsx"));
                #endregion
            }
        }
    }
}




