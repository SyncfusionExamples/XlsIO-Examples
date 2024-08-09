using System.IO;
using Syncfusion.XlsIO;

namespace Edit_Slicer
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //Access the table
                IListObject table = sheet.ListObjects[0];

                //Add slicer for the table
                sheet.Slicers.Add(table, 3, 11, 2);

                //Access the slicer
                ISlicer slicer = sheet.Slicers[0];

                //Slicer name
                slicer.Name = "Slicer1";

                //Slicer caption
                slicer.Caption = "Select any value";

                //Positioning a Slicer
                slicer.Top = 100;
                slicer.Left = 300;

                //Resize a Slicer
                slicer.Height = 200;
                slicer.Width = 150;

                //Resize Slicer item
                slicer.SlicerItemHeight = 0.4;
                slicer.SlicerItemWidth = 80;

                //Slicer columns
                slicer.NumberOfColumns = 2;

                //Slicer header
                slicer.DisplayHeader = true;

                //Slicer style
                slicer.SlicerStyle = ExcelSlicerStyle.SlicerStyleDark2;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("EditSlicer.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("EditSlicer.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
