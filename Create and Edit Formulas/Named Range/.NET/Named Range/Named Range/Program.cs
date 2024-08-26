using System.IO;
using Syncfusion.XlsIO;

namespace Named_Range
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
                IWorksheet sheet = workbook.Worksheets[0];
                sheet.Range["A1"].Value = "10";
                sheet.Range["B1"].Value = "20";

                //Defining a name in workbook level for the cell A1
                IName name1 = workbook.Names.Add("One");
                name1.RefersToRange = sheet.Range["A1"];

                //Defining a name in workbook level for the cell B1
                IName name2 = workbook.Names.Add("Two");
                name2.RefersToRange = sheet.Range["B1"];

                //Formula using defined names
                sheet.Range["C1"].Formula = "=SUM(One,Two)";

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Formula.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}




