using System.IO;
using Syncfusion.XlsIO;

namespace Group_Shapes
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                IShapes shapes = worksheet.Shapes;

                IShape[] groupItems;
                for (int i = 0; i < shapes.Count; i++)
                {
                    if (shapes[i].Name == "Development" || shapes[i].Name == "Production" || shapes[i].Name == "Sales")
                    {
                        groupItems = new IShape[] { shapes[i], shapes[i + 1], shapes[i + 2], shapes[i + 3], shapes[i + 4], shapes[i + 5] };
                        shapes.Group(groupItems);
                        i = -1;
                    }
                }

                groupItems = new IShape[] { shapes[0], shapes[1], shapes[2], shapes[3], shapes[4], shapes[5], shapes[6] };

                // Group the selected shapes
                shapes.Group(groupItems);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("GroupShapes.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}

