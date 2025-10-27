using System.IO;
using Syncfusion.XlsIO;

namespace Modify_Shape_Hyperlink
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Modify Shape Hyperlink
                //Modifying hyperlink’s screen tip through IWorksheet instance
                IHyperLink hyperlink = worksheet.HyperLinks[0];
                hyperlink.ScreenTip = "Syncfusion";

                //Modifying hyperlink’s screen tip through IShape instance
                hyperlink = worksheet.Shapes[1].Hyperlink;
                hyperlink.ScreenTip = "Mail Syncfusion";
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ModifyShapeHyperlink.xlsx"));
                #endregion
            }
        }
    }
}




