# How to assign data to multiple rows and columns in parallel without exceptions in C#?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to assign data to multiple rows and columns in parallel in C#.
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet worksheet = workbook.Worksheets[0];

    object m_lock = new object();
    int numberOfRows = 10;
    Parallel.For(1, numberOfRows, i =>
    {
        var rand = new Random();
        lock (m_lock)
        {
            worksheet.Range[i, 1].Value2 = string.Format("R{0}T{1}", i, rand.Next(10));
            worksheet.Range[i, 2].Value2 = string.Format("R{0}T{1}", i, rand.Next(10));                    
            worksheet.Range[i, 3].Value2 = DateTime.Now.AddDays(rand.NextDouble() * 10.0);
            worksheet.Range[i, 4].Value2 = DateTime.Now.AddDays(rand.NextDouble() * 10.0);
            worksheet.Range[i, 5].Value2 = rand.Next(2000);
            worksheet.Range[i, 6].Value2 = rand.Next(4000);                   
            worksheet.Range[i, 7].Value2 = rand.NextDouble() * 10000.0;
            worksheet.Range[i, 8].Value2 = rand.NextDouble() * 10000.0;
        }
    });

    #region Save
    //Saving the workbook
    workbook.SaveAs("../../../Output/Output.xlsx");
    #endregion
}
```