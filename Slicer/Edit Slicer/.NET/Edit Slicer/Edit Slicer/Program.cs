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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);
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
                workbook.SaveAs(Path.GetFullPath("Output/EditSlicer.xlsx"));
                #endregion
            }
        }
    }
}





