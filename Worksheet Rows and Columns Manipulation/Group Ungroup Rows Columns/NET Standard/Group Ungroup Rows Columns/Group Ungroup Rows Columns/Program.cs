using System.IO;
using Syncfusion.XlsIO;

namespace Group_Ungroup_Rows_Columns
{
    class Program
    {
        static void Main(string[] args)
        {
            GroupAndUnGroup obj = new GroupAndUnGroup();

            obj.GroupRowsColumns();
            obj.UngroupRowsColumns();
        }
    }
    public class GroupAndUnGroup
    {
        public void GroupRowsColumns()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../InputTemplate - ToGroup.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Group Rows
                //Group Rows
                worksheet.Range["A3:A7"].Group(ExcelGroupBy.ByRows, true);
                worksheet.Range["A11:A16"].Group(ExcelGroupBy.ByRows);
                #endregion

                #region Group Columns
                //Group Columns
                worksheet.Range["C1:D1"].Group(ExcelGroupBy.ByColumns, false);
                worksheet.Range["F1:G1"].Group(ExcelGroupBy.ByColumns);
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("GroupRowsColumns.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("GroupRowsColumns.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
        public void UngroupRowsColumns()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../InputTemplate - ToUngroup.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Un-Group Rows
                //Ungroup Rows
                worksheet.Range["A3:A7"].Ungroup(ExcelGroupBy.ByRows);
                #endregion

                #region Un-Group Columns
                //Ungroup Columns
                worksheet.Range["C1:D1"].Ungroup(ExcelGroupBy.ByColumns);
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("UngroupRowsColumns.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("UngroupRowsColumns.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
