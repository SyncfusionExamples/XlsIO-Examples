using System.IO;
using Syncfusion.XlsIO;

namespace Argument_Separator
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

                #region Set Separators
                //Setting the argument separator
                workbook.SetSeparators(';', ',');
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Formula.xlsx"));
                #endregion
            }
        }
    }
}




