using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.XlsIO;

namespace Nested_If_Function
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load the Excel document
                FileStream fileStream = new FileStream(@"../../Data/CustomFormula.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Get the used range for the worksheet
                IRange usedRange = worksheet.UsedRange;

                //Setting the nesdted if formula in the worksheet
                for (int row = 2; row <= usedRange.LastRow; row++)
                {
                    worksheet.Range["C" + row].Formula = "=IF(ISBLANK(A" + row + "), \"\", IF(OR(ISNUMBER(SEARCH(\"SMSF\", A" + row + ")), ISNUMBER(SEARCH(\"Trust\", A" + row + ")), ISNUMBER(SEARCH(\"Company\", A" + row + ")), ISNUMBER(SEARCH(\"PARTNERSHIP\", A" + row + "))), B" + row + ", \"\"))";
                }

                //Save the workbook as stream
                FileStream stream = new FileStream("Output.xlsx", FileMode.Create);
                workbook.SaveAs(stream);
            }
        }
    }
}
