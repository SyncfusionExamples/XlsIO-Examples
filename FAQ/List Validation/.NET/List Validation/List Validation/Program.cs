using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet["A1"].Value = "ListItem1";
            worksheet["A2"].Value = "ListItem2";
            worksheet["A3"].Value = "ListItem3";
            worksheet["A4"].Value = "ListItem4";
            
            worksheet.Range["C1"].Text = "Data Validation List in C3";
            worksheet.Range["C1"].AutofitColumns();

            //Data validation for the list
            IDataValidation listValidation = worksheet.Range["C3"].DataValidation;        
            listValidation.DataRange = worksheet.Range["A1:A4"];

            //Set the first item in the list as default value
            worksheet.Range["C3"].Value = worksheet.Range["C3"].DataValidation.DataRange.Cells[0].Value;
           

            #region Save
            //Saving the workbook
            workbook.SaveAs("../../../Output/ListValidation.xlsx");
            #endregion
        }
    }

    
}

