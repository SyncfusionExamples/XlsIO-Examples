using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace _3D_Chart
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(2);
                IWorksheet sheet = workbook.Worksheets[0];

                //Insert the data in sheet-1
                sheet.Range["B1"].Text = "Product-A";
                sheet.Range["C1"].Text = "Product-B";
                sheet.Range["D1"].Text = "Product-C";
                sheet.Range["A2"].Text = "Jan";
                sheet.Range["A3"].Text = "Feb";
                sheet.Range["B2"].Number = 25;
                sheet.Range["B3"].Number = 20;
                sheet.Range["C2"].Number = 35;
                sheet.Range["C3"].Number = 25;
                sheet.Range["D2"].Number = 40;
                sheet.Range["D3"].Number = 55;

                IChartShape chart = sheet.Charts.Add();
                chart.DataRange = sheet.Range["A1:D3"];
                chart.ChartType = ExcelChartType.Column_Clustered_3D;

                //Set Rotation of the 3D chart view
                chart.Rotation = 90;

                //Set Back wall fill option
                chart.BackWall.Fill.FillType = ExcelFillType.Gradient;
                //Set Back wall thickness
                chart.BackWall.Thickness = 10;
                //Set Texture Type
                chart.BackWall.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                chart.BackWall.Fill.GradientStyle = ExcelGradientStyle.Diagonl_Down;
                chart.BackWall.Fill.ForeColor = Color.WhiteSmoke;
                chart.BackWall.Fill.BackColor = Color.LightBlue;

                //Set side wall fill option
                chart.SideWall.Fill.FillType = ExcelFillType.SolidColor;
                //Set side wall fore and back color
                chart.SideWall.Fill.BackColor = Color.White;
                chart.SideWall.Fill.ForeColor = Color.White;

                //Set floor fill option
                chart.Floor.Fill.FillType = ExcelFillType.Pattern;
                chart.Floor.Fill.Pattern = ExcelGradientPattern.Pat_10_Percent;
                //Set floor fore and Back color
                chart.Floor.Fill.ForeColor = Color.Blue;
                chart.Floor.Fill.BackColor = Color.White;
                //Set floor thickness
                chart.Floor.Thickness = 3;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Chart.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Chart.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
