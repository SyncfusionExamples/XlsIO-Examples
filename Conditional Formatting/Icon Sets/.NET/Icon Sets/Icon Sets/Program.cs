using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Advanced_Conditional_Formats
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Create icon sets for the data in specified range
                IConditionalFormats conditionalFormats = worksheet.Range["E7:E46"].ConditionalFormats;
                IConditionalFormat conditionalFormat = conditionalFormats.AddCondition();
                conditionalFormat.FormatType = ExcelCFType.IconSet;
                IIconSet iconSet = conditionalFormat.IconSet;

                //Apply three symbols icon and hide the data in the specified range
                iconSet.IconSet = ExcelIconSetType.ThreeSymbols;
                iconSet.IconCriteria[1].Type = ConditionValueType.Percent;
                iconSet.IconCriteria[1].Value = "50";
                iconSet.IconCriteria[2].Type = ConditionValueType.Percent;
                iconSet.IconCriteria[2].Value = "75";
                iconSet.ShowIconOnly = true;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
                #endregion
            }
        }
    }
}





