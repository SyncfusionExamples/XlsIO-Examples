﻿using System.IO;
using Syncfusion.XlsIO;

namespace Icon_Filter
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Icon Filter
                //Creating an AutoFilter in the first worksheet. Specifying the AutoFilter range.
                worksheet.AutoFilters.FilterRange = worksheet.Range["A1:A8"];

                //Column index to which AutoFilter must be applied.
                IAutoFilter filter = worksheet.AutoFilters[0];

                //Applying Icon filter to filter based on applied icon set.
                filter.AddIconFilter(ExcelIconSetType.ThreeFlags, 2);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/IconFilter.xlsx"));
                #endregion
            }
        }
    }
}




