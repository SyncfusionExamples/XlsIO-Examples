using System;
using System.IO;
using Syncfusion.XlsIO;

namespace Precedent_and_Dependent_Cells
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
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
            }
        }
    }
}




