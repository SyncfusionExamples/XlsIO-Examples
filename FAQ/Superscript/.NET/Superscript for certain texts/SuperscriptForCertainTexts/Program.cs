using Syncfusion.XlsIO;

class Program
{        
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            //Create a workbook
            IWorkbook workbook = application.Workbooks.Open("../../../Data/Sample.xlsx");
            IWorksheet worksheet = workbook.Worksheets[0];

            //Add Text
            IRange range = worksheet.Range["A1"];
            IRichTextString richText = range.RichText;

            IFont superScript = workbook.CreateFont();
            superScript.Size = richText.GetFont(6).Size;
            superScript.FontName = richText.GetFont(6).FontName;
            superScript.Color = richText.GetFont(6).Color;
            superScript.Superscript = true;
            richText.SetFont(6, 6, superScript);


            //Save the workbook to disk in xlsx format
            workbook.SaveAs("../../../Output/Output.xlsx");
        }
    }

    
}
            

