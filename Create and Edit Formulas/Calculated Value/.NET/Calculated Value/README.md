# How to access the calculated value of a formula?

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

Step 5: Include the below code snippet in **Program.cs** to access the calculated value of a formula.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet sheet = workbook.Worksheets[0];

	sheet.Range["A1"].Value = "10";
	sheet.Range["B1"].Value = "20";

	sheet.Range["C1"].Formula = "=A1+B1";

	#region Calculated Value
	sheet.EnableSheetCalculations();

	//Returns the calculated value of a formula using the most current inputs
	string calculatedValue = sheet["C1"].CalculatedValue;
	sheet.Range["C3"].Value = "Calculated Value of the formula in C1 calculated through XlsIO is : " + calculatedValue;
	
	sheet.DisableSheetCalculations();
	#endregion

	sheet.Range["C3"].AutofitColumns();

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/Formula.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}