using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Forms;
using Syncfusion.XlsIO;
using System.Reflection;
using System.IO;

namespace Edit_Excel
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
        }
		void OnButtonClicked(object sender, EventArgs e)
		{
			using (ExcelEngine excelEngine = new ExcelEngine())
			{
				IApplication application = excelEngine.Excel;
				application.DefaultVersion = ExcelVersion.Xlsx;

				//"App" is the class of Portable project.
				Assembly assembly = typeof(App).GetTypeInfo().Assembly;
				Stream fileStream = assembly.GetManifestResourceStream("Edit_Excel.InputTemplate.xlsx");

				//Opens the workbook 
				IWorkbook workbook = application.Workbooks.Open(fileStream);

				//Access first worksheet from the workbook.
				IWorksheet worksheet = workbook.Worksheets[0];

				//Set Text in cell A3.
				worksheet.Range["A3"].Text = "Hello World";

				MemoryStream stream = new MemoryStream();
				workbook.SaveAs(stream);

				workbook.Close();

				//Save the stream into XLSX file
				Xamarin.Forms.DependencyService.Get<ISave>().SaveAndView("EditExcel.xlsx", "application/msexcel", stream);
			}
		}
    }
}
