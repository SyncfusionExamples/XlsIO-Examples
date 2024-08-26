using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Insert_Rows_with_Merged_Cells
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                FileStream inputStream = new FileStream("../../../InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                InsertWithMerge(3, 2, worksheet);
                                
                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/InsertRowswithMergedCells.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

            }
        }

        public static void InsertWithMerge(int rowIndex, int rowCount, IWorksheet worksheet)
        {
            while (rowCount > 0)
            {
                if (rowIndex > 1)
                {
                    worksheet.InsertRow(rowIndex, 1, ExcelInsertOptions.FormatAsBefore);

                    worksheet["A" + (rowIndex - 1)].EntireRow.CopyTo(worksheet["A" + rowIndex]);

                    worksheet["A" + rowIndex].EntireRow.Clear(ExcelClearOptions.ClearContent);
                }
                else
                    worksheet.InsertRow(rowIndex, 1, ExcelInsertOptions.FormatAsBefore);
                rowCount--;
            }
        }
    }
}




