# Set list data validation in an Excel document using C#

The Syncfusion&reg; [.NET Excel Library](https://www.syncfusion.com/document-processing/excel-framework/net/excel-library) (XlsIO) enables you to create, read, and edit Excel documents programmatically without Microsoft Excel or interop dependencies. Using this library, you can **set list data validation in an Excel document** using C#.

## Steps to set list data validation in an Excel document programmatically

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
```csharp
using System.IO;
using Syncfusion.XlsIO;
```
Step 5: Include the below code snippet in **Program.cs** to set list data validation in an Excel document.
```csharp
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
```

More information about setting list data validation in an Excel document can be found in this [documentation](https://help.syncfusion.com/document-processing/excel/excel-library/net/working-with-data-validation#list-validation) section.