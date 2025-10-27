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

                //Data validation for the list
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
                workbook.SaveAs(Path.GetFullPath("Output/ListValidation.xlsx"));
                #endregion
            }
        }
    }
}