using System.IO;
using Syncfusion.XlsIO;

namespace Shape_Hyperlink
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

                #region Shape Hyperlink
                //Adding hyperlink to TextBox 
                ITextBox textBox = worksheet.TextBoxes.AddTextBox(1, 1, 100, 100);
                IHyperLink hyperlink = worksheet.HyperLinks.Add((textBox as IShape), ExcelHyperLinkType.Url, "http://www.Syncfusion.com", "click here");

                //Adding hyperlink to AutoShape
                IShape autoShape = worksheet.Shapes.AddAutoShapes(AutoShapeType.Cloud, 10, 1, 100, 100);
                hyperlink = worksheet.HyperLinks.Add(autoShape, ExcelHyperLinkType.Url, "mailto:Username@syncfusion.com", "Send Mail");

                //Adding hyperlink to picture
                IPictureShape picture = worksheet.Pictures.AddPictureAsLink(5, 5, 10, 7, "../../../Image.png");
                hyperlink = worksheet.HyperLinks.Add(picture);
                hyperlink.Type = ExcelHyperLinkType.Unc;
                hyperlink.Address = "C://Documents and Settings";
                hyperlink.ScreenTip = "Click here for files";
				#endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ShapeHyperlink.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ShapeHyperlink.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
