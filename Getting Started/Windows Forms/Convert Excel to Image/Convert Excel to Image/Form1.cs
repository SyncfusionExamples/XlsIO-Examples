using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Convert_Excel_to_Image
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open("Sample.xlsx");
                IWorksheet worksheet = workbook.Worksheets[0];

                //Convert the Excel to Image
                Image image = worksheet.ConvertToImage(1, 1, 20, 4);

                //Save the image as jpeg
                image.Save("Sample.Jpeg", ImageFormat.Jpeg);
            }

            //Launch the Image file
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo = new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"../../Sample.Jpeg")) { UseShellExecute = true };
            process.Start();
        }
    }
}
