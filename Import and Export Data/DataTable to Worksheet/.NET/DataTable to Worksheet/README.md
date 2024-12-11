# How to import data from DataTable to Excel?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to import data from DataTable to Excel.
{% tabs %}
{% highlight c# tabtitle="C#" %}
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

            #region Import from Data Table
            //Initialize the DataTable
            DataTable table = SampleDataTable();
            //Import DataTable to the worksheet
            worksheet.ImportDataTable(table, true, 1, 1);
			#endregion

            #region Save
            //Saving the workbook
            FileStream outputStream = new FileStream(Path.GetFullPath("Output/ImportDataTable.xlsx"), FileMode.Create, FileAccess.Write);
            workbook.SaveAs(outputStream);
            #endregion

            //Dispose streams
            outputStream.Dispose();
        }
    }
    private static DataTable SampleDataTable()
    {
        //Create a DataTable with four columns
        DataTable table = new DataTable();
        table.Columns.Add("Dosage", typeof(int));
        table.Columns.Add("Drug", typeof(string));
        table.Columns.Add("Patient", typeof(string));
        table.Columns.Add("Date", typeof(DateTime));

        //Add five DataRows
        table.Rows.Add(25, "Indocin", "David", DateTime.Now);
        table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
        table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
        table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
        table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);

        return table;
    }
}
{% endhighlight %}
{% endtabs %}