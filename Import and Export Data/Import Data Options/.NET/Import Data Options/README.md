# How to import data from Collection Objects to Excel?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using System.IO;
using Syncfusion.XlsIO;
using System.Collections.Generic;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to import data from Collection Objects to Excel.
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

      //Import the data to worksheet with Import Data Options
      IList<Customer> reports = GetSalesReports();

      ExcelImportDataOptions importDataOptions = new ExcelImportDataOptions();
      importDataOptions.FirstRow = 2;
      importDataOptions.FirstColumn = 1;
      importDataOptions.IncludeHeader = false;
      importDataOptions.PreserveTypes = false;

      worksheet.ImportData(reports, importDataOptions);

      #region Save
      //Saving the workbook
      FileStream outputStream = new FileStream(Path.GetFullPath("Output/ImportDataOptions.xlsx"), FileMode.Create, FileAccess.Write);
      workbook.SaveAs(outputStream);
      #endregion

      //Dispose streams
      outputStream.Dispose();
    }
  }
  //Gets a list of sales reports
  public static List<Customer> GetSalesReports()
  {
    List<Customer> reports = new List<Customer>();
    reports.Add(new Customer("Andy Bernard", "45000", "58000"));
    reports.Add(new Customer("Jim Halpert", "34000", "65000"));
    reports.Add(new Customer("Karen Fillippelli", "75000", "64000"));
    reports.Add(new Customer("Phyllis Lapin", "56500", "33600"));
    reports.Add(new Customer("Stanley Hudson", "46500", "52000"));
    return reports;
  }
}

//Customer details
public class Customer
{
    public string SalesPerson { get; set; }
    public string SalesJanJun { get; set; }
    public string SalesJulDec { get; set; }

    public Customer(string name, string janToJun, string julToDec)
    {
        SalesPerson = name;
        SalesJanJun = janToJun;
        SalesJulDec = julToDec;
    }
}
{% endhighlight %}
{% endtabs %}