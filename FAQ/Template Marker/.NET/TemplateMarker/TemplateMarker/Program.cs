using Syncfusion.XlsIO;
using System.Data;


class Program
{
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));

            DataTable reports1 = new DataTable();
            reports1.Columns.Add("SalesPerson");
            reports1.Columns.Add("FromDate", typeof(DateTime));
            reports1.Columns.Add("ToDate", typeof(DateTime));
            reports1.Rows.Add("Andy Bernard", new DateTime(2014, 09, 08), new DateTime(2014, 09, 11));
            reports1.Rows.Add("Jim Halpert", new DateTime(2014, 09, 11), new DateTime(2014, 09, 15));

            //Create Template Marker Processor for Reports1
            ITemplateMarkersProcessor marker1 = workbook.CreateTemplateMarkersProcessor();
            //Add collection to marker variable
            marker1.AddVariable("Reports1", reports1, VariableTypeAction.None);
            //Process the markers in the template. And use UnknownVariableAction.Skip to skip the exception.
            marker1.ApplyMarkers(UnknownVariableAction.Skip);

            DataTable reports2 = new DataTable();
            reports2.Columns.Add("SalesPerson");
            reports2.Columns.Add("FromDate", typeof(DateTime));
            reports2.Columns.Add("ToDate", typeof(DateTime));
            reports2.Rows.Add("Karen Fillippelli", new DateTime(2014, 09, 15), new DateTime(2014, 09, 20));
            reports2.Rows.Add("Phyllis Lapin", new DateTime(2014, 09, 21), new DateTime(2014, 09, 25));
            reports2.Rows.Add("Stanley Hudson", new DateTime(2014, 09, 26), new DateTime(2014, 09, 30));

            //Create Template Marker Processor for Reports2
            ITemplateMarkersProcessor marker2 = workbook.CreateTemplateMarkersProcessor();
            //Add collection to marker variable
            marker2.AddVariable("Reports2", reports2, VariableTypeAction.None);
            //Process the markers in the template
            marker2.ApplyMarkers();

            //Saving the workbook
            workbook.Version = ExcelVersion.Xlsx;
            workbook.SaveAs(Path.GetFullPath(@"Output/TemplateMarker.xlsx"));
        }
    }
}