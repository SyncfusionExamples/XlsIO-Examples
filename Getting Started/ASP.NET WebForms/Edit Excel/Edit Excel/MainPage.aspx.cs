using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Syncfusion.XlsIO;

namespace Edit_Excel
{
    public partial class MainPage : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void OnButtonClicked(object sender, EventArgs e)
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Set the default application version
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load the existing Excel workbook into IWorkbook
                IWorkbook workbook = application.Workbooks.Open(Server.MapPath("App_Data/InputTemplate.xlsx"));

                //Get the first worksheet in the workbook into IWorksheet
                IWorksheet worksheet = workbook.Worksheets[0];

                //Assign some text in a cell
                worksheet.Range["A3"].Text = "Hello World";

                //Save the Excel document
                workbook.SaveAs("EditExcel.xlsx", Response, ExcelDownloadType.PromptDialog, ExcelHttpContentType.Excel2016);
            }
        }
    }
}