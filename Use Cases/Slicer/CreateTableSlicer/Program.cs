using System.IO;
using Syncfusion.XlsIO;

namespace CreateTableSlicer
{
    class Program
    {
        static void Main(string[] args)
        {
            // Initialize ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                // Set the default application version as Xlsx
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Open existing workbook with data
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);

                //Access first worksheet from the workbook.
                IWorksheet sheet = workbook.Worksheets[0];

                //Access the table.
                IListObject table = sheet.ListObjects[0];


                //Add Slicer to the Requester column(4th column) from the table at 11th row and 2nd column.
                sheet.Slicers.Add(table, 4, 11, 2);

                // Modify Slicer properties
                ISlicer slicer = sheet.Slicers[0];

                // Set Slicer caption, name, and size
                slicer.Caption = "Select Assignee";
                slicer.Name = "Assignees";
                slicer.Height = 200;
                slicer.Width = 200;

                //Apply built-in style for requester slicer
                slicer.SlicerStyle = ExcelSlicerStyle.SlicerStyleDark1;

                // Add Slicer to the Status column (5th column) from the table at 11th row and 4th column.
                sheet.Slicers.Add(table, 5, 11, 4);

                // Modify Slicer properties
                slicer = sheet.Slicers[1];
                slicer.Caption = "Select Status";
                slicer.Name = "Status";
                slicer.Height = 200;
                slicer.Width = 200;


                //Apply built-in style for status slicer
                slicer.SlicerStyle = ExcelSlicerStyle.SlicerStyleLight2;

                //Save the workbook
                workbook.SaveAs(Path.GetFullPath("Output/CreateTableSlicer.xlsx"));
            }
        }
    }
}
