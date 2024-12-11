# How to access the precedent and dependent cells in an Excel worksheet and workbook?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to access the precedent and dependent cells in an Excel worksheet and workbook.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
    IWorkbook workbook = application.Workbooks.Open(inputStream);
    IWorksheet worksheet = workbook.Worksheets[0];

    #region Precedents in Worksheet
    //Getting precedent cells from the worksheet
    IRange[] precedents_worksheet = worksheet["A1"].GetPrecedents();

    Console.WriteLine("Precedents of Sheet1!A1 in Worksheet are : " );
    foreach(IRange range in precedents_worksheet)
    {
        Console.WriteLine(range.Address);
    }
    Console.WriteLine();
    #endregion

    #region Precedents in Workbook
    //Getting precedent cells from the workbook
    IRange[] precedents_workbook = worksheet["A1"].GetPrecedents(true);

    Console.WriteLine("Precedents of Sheet1!A1 in Workbook are : ");
    foreach (IRange range in precedents_workbook)
    {
        Console.WriteLine(range.Address);
    }
    Console.WriteLine();
    #endregion

    #region Dependents in Worksheet
    //Getting dependent cells from the worksheet
    IRange[] dependents_worksheet = worksheet["C1"].GetDependents();

    Console.WriteLine("Dependents of Sheet1!C1 in Worksheet are : ");
    foreach (IRange range in dependents_worksheet)
    {
        Console.WriteLine(range.Address);
    }
    Console.WriteLine();
    #endregion

    #region Dependents in Workbook
    //Getting dependent cells from the workbook
    IRange[] dependents_workbook = worksheet["C1"].GetDependents(true);

    Console.WriteLine("Dependents of Sheet1!C1 in Workbook are : ");
    foreach (IRange range in dependents_workbook)
    {
        Console.WriteLine(range.Address);
    }
    Console.WriteLine();
    #endregion

    #region Direct Precedents in Worksheet
    //Getting precedent cells from the worksheet
    IRange[] direct_precedents_worksheet = worksheet["A1"].GetDirectPrecedents();

    Console.WriteLine("Direct Precedents of Sheet1!A1 in Worksheet are : ");
    foreach (IRange range in direct_precedents_worksheet)
    {
        Console.WriteLine(range.Address);
    }
    Console.WriteLine();
    #endregion

    #region Direct Precedents in Workbook
    //Getting precedent cells from the workbook
    IRange[] direct_precedents_workbook = worksheet["A1"].GetDirectPrecedents(true);

    Console.WriteLine("Direct Precedents of Sheet1!A1 in Workbook are : ");
    foreach (IRange range in direct_precedents_workbook)
    {
        Console.WriteLine(range.Address);
    }
    Console.WriteLine();
    #endregion

    #region Direct Dependents in Worksheet
    //Getting dependent cells from the worksheet
    IRange[] direct_dependents_worksheet = worksheet["C1"].GetDirectDependents();

    Console.WriteLine("Direct Dependents of Sheet1!C1 in Worksheet are : ");
    foreach (IRange range in direct_dependents_worksheet)
    {
        Console.WriteLine(range.Address);
    }
    Console.WriteLine();
    #endregion

    #region Direct Dependents in Workbook
    //Getting dependent cells from the workbook
    IRange[] direct_dependents_workbook = worksheet["C1"].GetDirectDependents(true);

    Console.WriteLine("Direct Dependents of Sheet1!C1 in Workbook are : ");
    foreach (IRange range in direct_dependents_workbook)
    {
        Console.WriteLine(range.Address);
    }
    Console.WriteLine();
    #endregion

    //Dispose streams
    inputStream.Dispose();
}
{% endhighlight %}
{% endtabs %} 