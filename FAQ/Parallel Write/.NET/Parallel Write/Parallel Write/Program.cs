using Syncfusion.XlsIO;

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
    }

    
}

