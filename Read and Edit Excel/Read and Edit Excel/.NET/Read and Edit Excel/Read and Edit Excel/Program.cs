﻿using Syncfusion.XlsIO;
using System.IO;

namespace Read_and_Edit_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new instance for ExcelEngine
            ExcelEngine excelEngine = new ExcelEngine();

            #region Open
            //Loads or open an existing workbook through Open method of IWorkbook
            FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
            IWorkbook workbook = excelEngine.Excel.Workbooks.Open(inputStream);
            #endregion

            //Set the version of the workbook
            workbook.Version = ExcelVersion.Xlsx;

            #region Edit
            //Set a value in Excel cell
            workbook.Worksheets[0].Range["A2"].Value = "Hello World";
            #endregion

            #region Save
            //Saving the workbook
            FileStream outputStream = new FileStream(Path.GetFullPath("Output/ReadandEditExcel.xlsx"), FileMode.Create, FileAccess.Write);
            workbook.SaveAs(outputStream);            
            #endregion

            #region Close
            //Close the instance of IWorkbook
            workbook.Close();
            #endregion

            //Dispose streams
            outputStream.Dispose();
            inputStream.Dispose();

            //Dispose the instance of ExcelEngine
            excelEngine.Dispose();
        }
    }
}





