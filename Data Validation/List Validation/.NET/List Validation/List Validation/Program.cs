using System.IO;
using Syncfusion.XlsIO;

namespace List_Validation
{
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
                FileStream outputStream = new FileStream("ListValidation.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ListValidation.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
