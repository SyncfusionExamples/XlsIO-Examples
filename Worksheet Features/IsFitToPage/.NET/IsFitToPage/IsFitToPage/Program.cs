using System.IO;
using Syncfusion.XlsIO;

namespace IsFitToPage
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
                // True to fit the content before printing
                sheet.PageSetup.IsFitToPage = true;

                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/IsFitToPage.xlsx"));
                #endregion
            }
        }
    }
}




