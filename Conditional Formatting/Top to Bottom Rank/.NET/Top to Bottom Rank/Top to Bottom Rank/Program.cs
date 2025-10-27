﻿using System.IO;
using Syncfusion.XlsIO;

namespace Top_to_Bottom_Rank
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

                //Applying conditional formatting to "N6:N35".
                IConditionalFormats formats = worksheet.Range["N6:N35"].ConditionalFormats;
                IConditionalFormat format = formats.AddCondition();

                //Applying top or bottom rule in the conditional formatting.
                format.FormatType = ExcelCFType.TopBottom;
                ITopBottom topBottom = format.TopBottom;

                //Set type as Top for TopBottom rule.
                topBottom.Type = ExcelCFTopBottomType.Top;

                //Set rank value for the TopBottom rule.
                topBottom.Rank = 10;

                //Set color for Conditional Formattting.
                format.BackColorRGB = Syncfusion.Drawing.Color.FromArgb(51, 153, 102);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/TopToBottomRank.xlsx"));
                #endregion
            }
        }
    }
}





