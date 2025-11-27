using Syncfusion.XlsIO;

class Program
{
    public static void Main()
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Create(3);
            IWorksheet dataSheet = workbook.Worksheets[0];
            IWorksheet pivotSheet1 = workbook.Worksheets[1];
            IWorksheet pivotSheet2 = workbook.Worksheets[2];

            //Add sample data
            dataSheet.Range["A1"].Text = "Product";
            dataSheet.Range["B1"].Text = "Sales";
            dataSheet.Range["C1"].Text = "Cost";

            dataSheet.Range["A2"].Text = "Laptop";
            dataSheet.Range["B2"].Number = 5000;
            dataSheet.Range["C2"].Number = 3000;

            dataSheet.Range["A3"].Text = "Tablet";
            dataSheet.Range["B3"].Number = 3000;
            dataSheet.Range["C3"].Number = 2000;

            dataSheet.Range["A4"].Text = "Phone";
            dataSheet.Range["B4"].Number = 4000;
            dataSheet.Range["C4"].Number = 2500;

            //CASE 1: Shared pivot cache — calculated field names must be unique
            IPivotCache sharedCache = workbook.PivotCaches.Add(dataSheet["A1:C4"]);

            IPivotTable pivot1 = pivotSheet1.PivotTables.Add("Pivot1", pivotSheet1["A1"], sharedCache);
            pivot1.Fields[0].Axis = PivotAxisTypes.Row;
            pivot1.DataFields.Add(pivot1.Fields[1], "Total Sales", PivotSubtotalTypes.Sum);
            pivot1.CalculatedFields.Add("Profit", "Sales - Cost");

            IPivotTable pivot2 = pivotSheet1.PivotTables.Add("Pivot2", pivotSheet1["F1"], sharedCache);
            pivot2.Fields[0].Axis = PivotAxisTypes.Row;
            pivot2.DataFields.Add(pivot2.Fields[2], "Total Cost", PivotSubtotalTypes.Sum);
            pivot2.CalculatedFields.Add("Margin", "Sales - Cost");

            //CASE 2: Separate pivot caches — identical or different calculated field names are allowed
            IPivotTable pivot3 = pivotSheet2.PivotTables.Add("Pivot3", pivotSheet2["A1"],
            workbook.PivotCaches.Add(dataSheet["A1:C4"]));
            pivot3.Fields[0].Axis = PivotAxisTypes.Row;
            pivot3.DataFields.Add(pivot3.Fields[1], "Total Sales", PivotSubtotalTypes.Sum);
            pivot3.CalculatedFields.Add("Profit", "Sales - Cost");

            IPivotTable pivot4 = pivotSheet2.PivotTables.Add("Pivot4", pivotSheet2["F1"],
            workbook.PivotCaches.Add(dataSheet["A1:C4"]));
            pivot4.Fields[0].Axis = PivotAxisTypes.Row;
            pivot4.DataFields.Add(pivot4.Fields[2], "Total Cost", PivotSubtotalTypes.Sum);
            pivot4.CalculatedFields.Add("Profit", "Sales - Cost");

            //Saving the workbook
            workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
        }
    }
}