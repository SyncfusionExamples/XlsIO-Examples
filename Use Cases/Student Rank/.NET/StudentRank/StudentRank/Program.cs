using Syncfusion.XlsIO;

class Program
{
    public static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.xlsx"), FileMode.Open, FileAccess.Read);
            IWorkbook workbook = application.Workbooks.Open(inputStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet["E1"].Text = "Rank";
            worksheet["E1"].CellStyle.Font.Bold = true;

            //Apply grading logic using nested IF formulas
            worksheet["E2"].Formula = "=IF(D2>=270,\"A\", IF(D2>=250,\"B\", IF(D2>=230,\"C\",\"D\")))";
            worksheet["E3"].Formula = "=IF(D3>=270,\"A\", IF(D3>=250,\"B\", IF(D3>=230,\"C\",\"D\")))";
            worksheet["E4"].Formula = "=IF(D4>=270,\"A\", IF(D4>=250,\"B\", IF(D4>=230,\"C\",\"D\")))";
            worksheet["E5"].Formula = "=IF(D5>=270,\"A\", IF(D5>=250,\"B\", IF(D5>=230,\"C\",\"D\")))";
            worksheet["E6"].Formula = "=IF(D6>=270,\"A\", IF(D6>=250,\"B\", IF(D6>=230,\"C\",\"D\")))";
            worksheet["E7"].Formula = "=IF(D7>=270,\"A\", IF(D7>=250,\"B\", IF(D7>=230,\"C\",\"D\")))";

            //Align given range to the right
            worksheet.Range["E2:E7"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;

            //Saving the workbook 
            FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.ReadWrite);
            workbook.SaveAs(outputStream);

            //Dispose streams
            outputStream.Dispose();
            inputStream.Dispose();
        }
    }
}