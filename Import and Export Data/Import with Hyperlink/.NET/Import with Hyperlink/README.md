# How to import data from collection objects with hyperlinks into Excel?

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

Step 5: Include the below code snippet in **Program.cs** to import data from collection objects with hyperlinks into Excel.
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

      //Import the data to worksheet
      IList<Company> reports = GetCompanyDetails();
      worksheet.ImportData(reports, 2, 1, false);

      #region Save
      //Saving the workbook
      FileStream outputStream = new FileStream(Path.GetFullPath("Output/ImportData.xlsx"), FileMode.Create, FileAccess.Write);
      workbook.SaveAs(outputStream);
      #endregion

      //Dispose streams
      outputStream.Dispose();
    }
  }
  //Gets a list of company details
  private static List<Company> GetCompanyDetails()
  {
    List<Company> companyList = new List<Company>();

    Company company = new Company();
    company.Name = "Syncfusion";
    Hyperlink link = new Hyperlink("https://www.syncfusion.com", "", "", "Syncfusion", ExcelHyperLinkType.Url, null);
    company.Link = link;
    companyList.Add(company);

    company = new Company();
    company.Name = "Microsoft";
    link = new Hyperlink("https://www.microsoft.com", "", "", "Microsoft", ExcelHyperLinkType.Url, null);
    company.Link = link;
    companyList.Add(company);

    company = new Company();
    company.Name = "Google";
    link = new Hyperlink("https://www.google.com", "", "", "Google", ExcelHyperLinkType.Url, null);
    company.Link = link;
    companyList.Add(company);

    return companyList;
  }
}
public class Hyperlink : IHyperLink
{
  public IApplication Application { get; }
  public object Parent { get; }
  public string Address { get; set; }
  public string Name { get; }
  public IRange Range { get; }
  public string ScreenTip { get; set; }
  public string SubAddress { get; set; }
  public string TextToDisplay { get; set; }
  public ExcelHyperLinkType Type { get; set; }
  public IShape Shape { get; }
  public ExcelHyperlinkAttachedType AttachedType { get; }
  public byte[] Image { get; set; }

  public Hyperlink(string address, string subAddress, string screenTip, string textToDisplay, ExcelHyperLinkType type, byte[] image)
  {
    Address = address;
    ScreenTip = screenTip;
    SubAddress = subAddress;
    TextToDisplay = textToDisplay;
    Type = type;
    Image = image;
  }
}

public class Company
{
  public string Name { get; set; }
  public Hyperlink Link { get; set; }
}
      
{% endhighlight %}
{% endtabs %}