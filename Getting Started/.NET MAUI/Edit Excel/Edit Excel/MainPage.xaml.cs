using System;
using Microsoft.Maui.Controls;
using Microsoft.Maui.Essentials;
using Syncfusion.XlsIO;
using System.IO;
using System.Reflection;

namespace Edit_Excel
{
	public partial class MainPage : ContentPage
	{
		public MainPage()
		{
			InitializeComponent();
		}

		private void editExcel_Click(object sender, EventArgs e)
		{
            //Create an instance of ExcelEngine.
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                Assembly executingAssembly = typeof(App).GetTypeInfo().Assembly;
                Stream inputStream = executingAssembly.GetManifestResourceStream("Edit_Excel.InputTemplate.xlsx");

                //Create a workbook with a worksheet
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Access first worksheet from the workbook instance.
                IWorksheet worksheet = workbook.Worksheets[0];

                //Set Text in cell A3.
                worksheet.Range["A3"].Text = "Hello World";

                MemoryStream ms = new MemoryStream();
                workbook.SaveAs(ms);
                ms.Position = 0;

                //Saves the memory stream as a file.
                DependencyService.Get<ISave>().SaveAndView("EditExcel.xlsx", "application/excel", ms);
            }
        }
	}
}
