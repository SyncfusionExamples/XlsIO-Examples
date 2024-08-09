using System;
using System.IO;
using Syncfusion.XlsIO;

namespace Number_Format
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

                worksheet.Range["A1"].Text = "DATA";
                worksheet.Range["B1"].Text = "NUMBER FORMAT APPLIED";
                worksheet.Range["C1"].Text = "RESULT";
                IStyle headingStyle = workbook.Styles.Add("HeadingStyle");
                headingStyle.Font.Bold = true;
                headingStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                worksheet.Range["A1:C1"].CellStyle = headingStyle;

                #region Applying Number Format
                //Applying different number formats
                worksheet.Range["A2"].Text = "1000000.00075";
                worksheet.Range["B2"].Text = "0.00";
                worksheet.Range["C2"].NumberFormat = "0.00";
                worksheet.Range["C2"].Number = 1000000.00075;
                worksheet.Range["A3"].Text = "1000000.500";
                worksheet.Range["B3"].Text = "###,##";
                worksheet.Range["C3"].NumberFormat = "###,##";
                worksheet.Range["C3"].Number = 1000000.500;
                worksheet.Range["A5"].Text = "10000";
                worksheet.Range["B5"].Text = "0.00";
                worksheet.Range["C5"].NumberFormat = "0.00";
                worksheet.Range["C5"].Number = 10000;
                worksheet.Range["A6"].Text = "-500";
                worksheet.Range["B6"].Text = "[Blue]#,##0";
                worksheet.Range["C6"].NumberFormat = "[Blue]#,##0";
                worksheet.Range["C6"].Number = -500;
                worksheet.Range["A7"].Text = "0.000000000000000000001234567890";
                worksheet.Range["B7"].Text = "0.000000000000000000000000000000";
                worksheet.Range["C7"].NumberFormat = "0.000000000000000000000000000000";
                worksheet.Range["C7"].Number = 0.000000000000000000001234567890;
                worksheet.Range["A9"].Text = "1.20";
                worksheet.Range["B9"].Text = "0.00E+00";
                worksheet.Range["C9"].NumberFormat = "0.00E+00";
                worksheet.Range["C9"].Number = 1.20;

                //Applying percentage format
                worksheet.Range["A10"].Text = "1.20";
                worksheet.Range["B10"].Text = "0.00%";
                worksheet.Range["C10"].NumberFormat = "0.00%";
                worksheet.Range["C10"].Number = 1.20;

                //Applying date format
                worksheet.Range["A11"].Text = new DateTime(2005, 12, 25).ToString();
                worksheet.Range["B11"].Text = "m/d/yyyy";
                worksheet.Range["C11"].NumberFormat = "m/d/yyyy";
                worksheet.Range["C11"].DateTime = new DateTime(2005, 12, 25);

                //Applying currency format
                worksheet.Range["A12"].Text = "1.20";
                worksheet.Range["B12"].Text = "$#,##0.00";
                worksheet.Range["C12"].NumberFormat = "$#,##0.00";
                worksheet.Range["C12"].Number = 1.20;

                //Applying accounting format
                worksheet.Range["A12"].Text = "234";
                worksheet.Range["B12"].Text = "_($* #,##0_)";
                worksheet.Range["C12"].NumberFormat = "_($* #,##0_)";
                worksheet.Range["C12"].Number = 234;
                #endregion

                #region Accessing Value with Number Format
                //Get display text of the cell
                string text = worksheet.Range["C12"].DisplayText;
                #endregion

                //Fit column width to data
                worksheet.UsedRange.AutofitColumns();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("NumberFormat.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("NumberFormat.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
