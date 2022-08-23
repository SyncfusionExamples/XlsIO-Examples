using System;
using System.Data;
using System.Windows.Forms;
using Syncfusion.XlsIO;

namespace DataGridView_to_Worksheet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnCreate_Click(object sender, EventArgs e)
        {
            //Initialize the Excel Engine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Initialize DataGridView control
                DataGridView dataGridView = new DataGridView();

                //Assign data source
                dataGridView.DataSource = GetDataTable();

                //Add control to group box
                groupBox.Controls.Add(dataGridView);

                //Assign the datagridview size
                dataGridView.Size = new System.Drawing.Size(441, 150);

                //Initialize Application
                IApplication application = excelEngine.Excel;

                //Set default version for application
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Create a new workbook
                IWorkbook workbook = application.Workbooks.Create(1);

                //Accessing first worksheet in the workbook
                IWorksheet worksheet = workbook.Worksheets[0];

                //Import data from DataGridView control
                worksheet.ImportDataGridView(dataGridView, 1, 1, true, true);

                //Save the workbook
                workbook.SaveAs("Output.xlsx");
                System.Diagnostics.Process.Start("Output.xlsx");
            }
        }
        private static DataTable GetDataTable()
        {
            Random r = new Random();
            DataTable dt = new DataTable("NumbersTable");

            int nCols = 4;
            int nRows = 10;

            for (int i = 0; i < nCols; i++)
                dt.Columns.Add(new DataColumn("Column" + i.ToString()));

            for (int i = 0; i < nRows; ++i)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < nCols; j++)
                    dr[j] = r.Next(0, 10);
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}
