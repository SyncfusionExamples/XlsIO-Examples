using Syncfusion.XlsIO;

namespace RemoveCorruptedImages
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));
                foreach (IWorksheet sheet in workbook.Worksheets)
                {
                    for (int i = 0; i < sheet.Pictures.Count; i++)
                    {
                        if (sheet.Pictures[i].Picture.ImageData.Length <= 0)
                        {
                            // Remove the corrupted image.
                            Console.WriteLine("Image removed due to corruption");
                            sheet.Pictures[i].Remove();
                        }
                    }
                }

                //Save the workbook
                workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
            }
        }
    }
}