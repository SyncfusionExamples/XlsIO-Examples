using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Syncfusion.XlsIO;

namespace Edit_Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public void btnEdit_Click(object sender, EventArgs e)
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Set the default application version
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load the existing Excel workbook into IWorkbook
                IWorkbook workbook = application.Workbooks.Open("../../InputTemplate.xlsx");

                //Get the first worksheet in the workbook into IWorksheet
                IWorksheet worksheet = workbook.Worksheets[0];

                //Assign some text in a cell
                worksheet.Range["A3"].Text = "Hello World";

                //Save the Excel document
                workbook.SaveAs("EditExcel.xlsx");
                System.Diagnostics.Process.Start("EditExcel.xlsx");
            }
        }
    }
}
