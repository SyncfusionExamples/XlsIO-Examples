# How to export data from Excel into collection objects?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using System.IO;
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to export data from Excel into collection objects.
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
            FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
            IWorkbook workbook = application.Workbooks.Open(inputStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Export worksheet data into Collection Objects
            List<Report> collectionObjects = worksheet.ExportData<Report>(1, 1, 10, 3);

            //Dispose streams
            inputStream.Dispose();
        }
    }
    public class Report
    {
        [DisplayNameAttribute("Sales Person Name")]
        public string SalesPerson { get; set; }
        [Bindable(false)]
        public string SalesJanJun { get; set; }
        public string SalesJulDec { get; set; }

        public Report()
        {

        }
    }
}
{% endhighlight %}
{% endtabs %}