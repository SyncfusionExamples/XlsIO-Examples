# How to expand or collapse groups in the worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to expand or collapse groups in the worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
class Program
{
    static void Main(string[] args)
    {
        ExpandandCollapse obj = new ExpandandCollapse();

        obj.ExpandGroups();
        obj.CollapseGroups();
    }
}
public class ExpandandCollapse
{
    public void ExpandGroups()
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate - To Expand.xlsx"), FileMode.Open, FileAccess.Read);
            IWorkbook workbook = application.Workbooks.Open(inputStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            #region Expand Groups
            //Expand row groups
            worksheet.Range["A3:A7"].ExpandGroup(ExcelGroupBy.ByRows, ExpandCollapseFlags.ExpandParent);
            worksheet.Range["A11:A16"].ExpandGroup(ExcelGroupBy.ByRows);

            //Expand column groups
            worksheet.Range["C1:D1"].ExpandGroup(ExcelGroupBy.ByColumns, ExpandCollapseFlags.ExpandParent);
            worksheet.Range["F1:G1"].ExpandGroup(ExcelGroupBy.ByColumns);
            #endregion

            #region Save
            //Saving the workbook
            FileStream outputStream = new FileStream(Path.GetFullPath("Output/ExpandGroups.xlsx"), FileMode.Create, FileAccess.Write);
            workbook.SaveAs(outputStream);
            #endregion

            //Dispose streams
            outputStream.Dispose();
            inputStream.Dispose();
        }
    }
    public void CollapseGroups()
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate - To Collapse.xlsx"), FileMode.Open, FileAccess.Read);
            IWorkbook workbook = application.Workbooks.Open(inputStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            #region Collapse Groups
            //Collapse row groups
            worksheet.Range["A3:A7"].CollapseGroup(ExcelGroupBy.ByRows);
            worksheet.Range["A11:A16"].CollapseGroup(ExcelGroupBy.ByRows);

            //Collapse column groups
            worksheet.Range["C1:D1"].CollapseGroup(ExcelGroupBy.ByColumns);
            worksheet.Range["F1:G1"].CollapseGroup(ExcelGroupBy.ByColumns);
            #endregion

            #region Save
            //Saving the workbook
            FileStream outputStream = new FileStream(Path.GetFullPath("Output/CollapseGroups.xlsx"), FileMode.Create, FileAccess.Write);
            workbook.SaveAs(outputStream);
            #endregion

            //Dispose streams
            outputStream.Dispose();
            inputStream.Dispose();
        }
    }
}
{% endhighlight %}
{% endtabs %}