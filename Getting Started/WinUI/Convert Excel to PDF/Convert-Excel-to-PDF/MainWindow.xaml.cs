// Copyright (c) Microsoft Corporation and Contributors.
// Licensed under the MIT License.

using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace Convert_Excel_to_PDF
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        public MainWindow()
        {
            this.InitializeComponent();
        }
        private void ConvertExceltoPDF(object sender, RoutedEventArgs e)
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load an existing file
                Assembly assembly = typeof(App).GetTypeInfo().Assembly;
                using (Stream inputStream = assembly.GetManifestResourceStream("Convert_Excel_to_PDF.InputTemplate.xlsx"))
                {
                    IWorkbook workbook = application.Workbooks.Open(inputStream);

                    //Initialize XlsIO renderer.
                    XlsIORenderer renderer = new XlsIORenderer();

                    //Convert Excel document into PDF document 
                    PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                    //Create the MemoryStream to save the converted PDF.      
                    MemoryStream pdfStream = new MemoryStream();

                    //Save the converted PDF document to MemoryStream.
                    pdfDocument.Save(pdfStream);
                    pdfStream.Position = 0;
                 
                    // Save the PDF file or perform any other action with the PDF
                    SaveHelper.SaveAndLaunch("Sample.pdf", pdfStream); 
                    
                }
            }
        }
    }
}
