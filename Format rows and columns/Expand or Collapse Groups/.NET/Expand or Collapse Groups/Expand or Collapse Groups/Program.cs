using System.IO;
using Syncfusion.XlsIO;

namespace Expand_or_Collapse_Groups
{
    class Program
    {
        static void Main(string[] args)
        {
            ExpandandCollapse obj = new ExpandandCollapse();

            obj.ExpandGroups();
            obj.CollapseGroups();
        }
    }
    public class ExpandandCollapse
    {
        public void ExpandGroups()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate - To Expand.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Expand Groups
                //Expand row groups
                worksheet.Range["A3:A7"].ExpandGroup(ExcelGroupBy.ByRows, ExpandCollapseFlags.ExpandParent);
                worksheet.Range["A11:A16"].ExpandGroup(ExcelGroupBy.ByRows);

                //Expand column groups
                worksheet.Range["C1:D1"].ExpandGroup(ExcelGroupBy.ByColumns, ExpandCollapseFlags.ExpandParent);
                worksheet.Range["F1:G1"].ExpandGroup(ExcelGroupBy.ByColumns);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ExpandGroups.xlsx"));
                #endregion
            }
        }
        public void CollapseGroups()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate - To Collapse.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Collapse Groups
                //Collapse row groups
                worksheet.Range["A3:A7"].CollapseGroup(ExcelGroupBy.ByRows);
                worksheet.Range["A11:A16"].CollapseGroup(ExcelGroupBy.ByRows);

                //Collapse column groups
                worksheet.Range["C1:D1"].CollapseGroup(ExcelGroupBy.ByColumns);
                worksheet.Range["F1:G1"].CollapseGroup(ExcelGroupBy.ByColumns);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/CollapseGroups.xlsx"));
                #endregion
            }
        }
    }
}





