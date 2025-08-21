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
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
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
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}





