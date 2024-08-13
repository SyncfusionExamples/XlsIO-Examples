using System.IO;
using Syncfusion.XlsIO;

namespace Excel_to_ODS
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

                worksheet.Range["A1"].Text = "Month";
                worksheet.Range["B1"].Text = "Sales";
                worksheet.Range["A5"].Text = "Total";
                worksheet.Range["A2"].Text = "January";
                worksheet.Range["A3"].Text = "February";

                worksheet.AutofitColumn(1);

                worksheet.Range["B2"].Number = 68878;
                worksheet.Range["B3"].Number = 71550;
                worksheet.Range["B5"].Formula = "SUM(B2:B4)";

                //Comments
                IComment comment = worksheet.Range["B5"].AddComment();
                comment.RichText.Text = "This cell has formula.";

                IRichTextString richText = comment.RichText;

                IFont blueFont = workbook.CreateFont();
                blueFont.Color = ExcelKnownColors.Blue;
                richText.SetFont(0, 13, blueFont);

                IFont redFont = workbook.CreateFont();
                redFont.Color = ExcelKnownColors.Red;
                richText.SetFont(14, 20, redFont);

                //Formatting
                IStyle style = workbook.Styles.Add("Style1");
                style.Color = Syncfusion.Drawing.Color.DarkBlue;
                style.Font.Color = ExcelKnownColors.WhiteCustom;

                worksheet.Range["A1:B1"].CellStyleName = "Style1";
                worksheet.Range["A5:B5"].CellStyleName = "Style1";

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("ExcelToODS.ods", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream, ExcelSaveType.SaveAsODS);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ExcelToODS.ods")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
