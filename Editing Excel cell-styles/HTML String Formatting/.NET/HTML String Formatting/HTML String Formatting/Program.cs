// See https://aka.ms/new-console-template for more information
using Syncfusion.XlsIO;
using static System.Net.Mime.MediaTypeNames;

using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet worksheet = workbook.Worksheets[0];

    //Add HTML string
    worksheet.Range["A1"].HtmlString = "<font style=\"color:red;font-family:Magneto;font-size:12px; \">Welcome Syncfusion</font>";

    //Assign HTML string as text to different range
    worksheet.Range["A2"].Text = worksheet.Range["A1"].HtmlString;

    #region Save
    //Saving the workbook
    FileStream outputStream = new FileStream("HTMLString.xlsx", FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);
    #endregion

    //Dispose streams
    outputStream.Dispose();

    System.Diagnostics.Process process = new System.Diagnostics.Process();
    process.StartInfo = new System.Diagnostics.ProcessStartInfo("HTMLString.xlsx")
    {
        UseShellExecute = true
    };
    process.Start();
}
