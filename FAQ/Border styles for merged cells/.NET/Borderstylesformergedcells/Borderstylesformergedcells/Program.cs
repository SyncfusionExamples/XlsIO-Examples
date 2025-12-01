using Syncfusion.XlsIO;

namespace Sample
{   
    class Program
    {        
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open("../../../Data/Input.xlsx");
                IWorksheet worksheet = workbook.Worksheets[0];

                # region Creating a new style
                IStyle style = workbook.Styles.Add("NewStyle");

                style.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thick;
                style.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thick;
                style.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thick;
                style.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thick;
                style.Borders.Color = ExcelKnownColors.Red;

                style.Font.Bold = true;
                style.Font.Color = ExcelKnownColors.Green;
                style.Font.Size = 24;
                #endregion

                //Applying style for merged region
                worksheet.Range[2, 1].MergeArea.CellStyle = style;

                workbook.SaveAs("../../../Output/MergeArea_Style.xlsx");
            }
        }

        
    }
            
}
