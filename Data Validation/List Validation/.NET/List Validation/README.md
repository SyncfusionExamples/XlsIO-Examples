# How to set list data validation in Excel?

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

Step 5: Include the below code snippet in **Program.cs** to set list data validation in Excel.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet worksheet = workbook.Worksheets[0];

    //Data Validation for List
    IDataValidation listValidation = worksheet.Range["C3"].DataValidation;
    worksheet.Range["C1"].Text = "Data Validation List in C3";
    worksheet.Range["C1"].AutofitColumns();
    listValidation.ListOfValues = new string[] { "ListItem1", "ListItem2", "ListItem3" };

    //Shows the error message
    listValidation.ErrorBoxText = "Choose the value from the list";
    listValidation.ErrorBoxTitle = "ERROR";
    listValidation.PromptBoxText = "Data validation for list";
    listValidation.IsPromptBoxVisible = true;
    listValidation.ShowPromptBox = true;

    #region Save
    //Saving the workbook
    FileStream outputStream = new FileStream(Path.GetFullPath("Output/ListValidation.xlsx"), FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);
    #endregion

    //Dispose streams
    outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}