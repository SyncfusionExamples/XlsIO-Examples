using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using Syncfusion.XlsIO;
using System.Reflection;
using Windows.Storage;
using Windows.Storage.Pickers;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Edit_Excel
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
        }

        async void OnButtonClicked(object sender, RoutedEventArgs e)
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Opens the workbook. 
                Assembly executingAssembly = typeof(MainPage).GetTypeInfo().Assembly;
                Stream inputStream = executingAssembly.GetManifestResourceStream("Edit_Excel.InputTemplate.xlsx");
                IWorkbook workbook = await application.Workbooks.OpenAsync(inputStream);

                //Access first worksheet from the workbook.
                IWorksheet worksheet = workbook.Worksheets[0];

                //Set Text in cell A3.
                worksheet.Range["A3"].Text = "Hello World";

                // Save the Workbook
                StorageFile storageFile;
                if (!(Windows.Foundation.Metadata.ApiInformation.IsTypePresent("Windows.Phone.UI.Input.HardwareButtons")))
                {
                    FileSavePicker savePicker = new FileSavePicker();
                    savePicker.SuggestedStartLocation = PickerLocationId.Desktop;
                    savePicker.SuggestedFileName = "Edit Excel";
                    savePicker.FileTypeChoices.Add("Excel Files", new List<string>() { ".xlsx" });
                    storageFile = await savePicker.PickSaveFileAsync();
                }
                else
                {
                    StorageFolder local = Windows.Storage.ApplicationData.Current.LocalFolder;
                    storageFile = await local.CreateFileAsync("Edit Excel.xlsx", CreationCollisionOption.ReplaceExisting);
                }
                //Saving the workbook
                await workbook.SaveAsAsync(storageFile);

                // Launch the saved file
                await Windows.System.Launcher.LaunchFileAsync(storageFile);
            }
        }
    }
}
