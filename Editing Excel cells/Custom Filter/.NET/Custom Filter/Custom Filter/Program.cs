﻿using System.IO;
using Syncfusion.XlsIO;

namespace Custom_Filter
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

                #region Custom Filter
                //Creating an AutoFilter in the first worksheet. Specifying the AutoFilter range
                worksheet.AutoFilters.FilterRange = worksheet.Range["A1:A11"];
                IAutoFilter filter = worksheet.AutoFilters[0];

                //Specifying first condition
                IAutoFilterCondition firstCondition = filter.FirstCondition;
                firstCondition.ConditionOperator = ExcelFilterCondition.Greater;
                firstCondition.Double = 100;

                //Specifying second condition
                IAutoFilterCondition secondCondition = filter.SecondCondition;
                secondCondition.ConditionOperator = ExcelFilterCondition.Less;
                secondCondition.Double = 200;
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/CustomFilter.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}




