using Syncfusion.Drawing;
using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Apply_Custom_Style
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

                //Create data for table
                worksheet[1, 1].Text = "Products";
                worksheet[1, 2].Text = "Qtr1";
                worksheet[1, 3].Text = "Qtr2";
                worksheet[1, 4].Text = "Qtr3";
                worksheet[1, 5].Text = "Qtr4";

                worksheet[2, 1].Text = "Alfreds Futterkiste";
                worksheet[2, 2].Number = 744.6;
                worksheet[2, 3].Number = 162.56;
                worksheet[2, 4].Number = 5079.6;
                worksheet[2, 5].Number = 1249.2;

                worksheet[3, 1].Text = "Antonio Moreno";
                worksheet[3, 2].Number = 5079.6;
                worksheet[3, 3].Number = 1249.2;
                worksheet[3, 4].Number = 943.89;
                worksheet[3, 5].Number = 349.6;

                worksheet[4, 1].Text = "Around the Horn";
                worksheet[4, 2].Number = 1267.5;
                worksheet[4, 3].Number = 1062.5;
                worksheet[4, 4].Number = 744.6;
                worksheet[4, 5].Number = 162.56;

                worksheet[5, 1].Text = "Bon app";
                worksheet[5, 2].Number = 1418;
                worksheet[5, 3].Number = 756;
                worksheet[5, 4].Number = 1267.5;
                worksheet[5, 5].Number = 1062.5;

                worksheet[6, 1].Text = "Eastern Connection";
                worksheet[6, 2].Number = 4728;
                worksheet[6, 3].Number = 4547.92;
                worksheet[6, 4].Number = 1418;
                worksheet[6, 5].Number = 756;

                worksheet[7, 1].Text = "Ernst Handel";
                worksheet[7, 2].Number = 943.89;
                worksheet[7, 3].Number = 349.6;
                worksheet[7, 4].Number = 4728;
                worksheet[7, 5].Number = 4547.92;

                //Create style for table number format
                IStyle style = workbook.Styles.Add("CurrencyFormat");
                style.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* \" - \"??_);_(@_)";
                worksheet["B2:E8"].CellStyleName = "CurrencyFormat";

                //Create table
                IListObject table = worksheet.ListObjects.Create("Table1", worksheet["A1:E7"]);

                //Apply custom table style
                ITableStyles tableStyles = workbook.TableStyles;
                ITableStyle tableStyle = tableStyles.Add("Table Style 1");
                ITableStyleElements tableStyleElements = tableStyle.TableStyleElements;
                ITableStyleElement tableStyleElement = tableStyleElements.Add(ExcelTableStyleElementType.SecondColumnStripe);
                tableStyleElement.BackColorRGB = Color.FromArgb(217, 225, 242);

                ITableStyleElement tableStyleElement1 = tableStyleElements.Add(ExcelTableStyleElementType.FirstColumn);
                tableStyleElement1.FontColorRGB = Color.FromArgb(128, 128, 128);

                ITableStyleElement tableStyleElement2 = tableStyleElements.Add(ExcelTableStyleElementType.HeaderRow);
                tableStyleElement2.FontColor = ExcelKnownColors.White;
                tableStyleElement2.BackColorRGB = Color.FromArgb(0, 112, 192);

                ITableStyleElement tableStyleElement3 = tableStyleElements.Add(ExcelTableStyleElementType.TotalRow);
                tableStyleElement3.BackColorRGB = Color.FromArgb(0, 112, 192);
                tableStyleElement3.FontColor = ExcelKnownColors.White;

                table.TableStyleName = tableStyle.Name;

                //Total row
                table.ShowTotals = true;
                table.ShowFirstColumn = true;
                table.ShowTableStyleColumnStripes = true;
                table.ShowTableStyleRowStripes = true;
                table.Columns[0].TotalsRowLabel = "Total";
                table.Columns[1].TotalsCalculation = ExcelTotalsCalculation.Sum;
                table.Columns[2].TotalsCalculation = ExcelTotalsCalculation.Sum;
                table.Columns[3].TotalsCalculation = ExcelTotalsCalculation.Sum;
                table.Columns[4].TotalsCalculation = ExcelTotalsCalculation.Sum;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/CustomTableStyle.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("CustomTableStyle.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
