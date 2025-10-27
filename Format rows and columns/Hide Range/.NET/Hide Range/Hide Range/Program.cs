using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Collections;

namespace Hide_Range
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

                IRange range = worksheet.Range["D4"];

                #region Hide single cell
                //Hiding the range ‘D4’
                worksheet.ShowRange(range, false);
                #endregion

                IRange firstRange = worksheet.Range["F6:I9"];
                IRange secondRange = worksheet.Range["C15:G20"];
                RangesCollection rangeCollection = new RangesCollection(application, worksheet);
                rangeCollection.Add(firstRange);
                rangeCollection.Add(secondRange);

                #region Hide multiple cells
                //Hiding a collection of ranges
                worksheet.ShowRange(rangeCollection, false);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/HideRange.xlsx"));
                #endregion
            }
        }
    }
}




