using System.IO;
using Syncfusion.XlsIO;

namespace Create_Style
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

                #region Create Style
                //Creating a new style with cell back color, fill pattern and font attribute
                IStyle style = workbook.Styles.Add("NewStyle");
                style.Color = Syncfusion.Drawing.Color.LightGreen;
                style.FillPattern = ExcelPattern.DarkUpwardDiagonal;
                style.Font.Bold = true;
                worksheet.Range["B2"].CellStyle = style;
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/CreateStyle.xlsx"));
                #endregion
            }
        }
    }
}




