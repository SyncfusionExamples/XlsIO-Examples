using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.IO;

namespace Import_with_Hyperlinks
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Create Template Marker Processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                //Add collection to the marker variables where the name should match with input template
                marker.AddVariable("Company", GetCompanyDetails());

                //Process the markers in the template
                marker.ApplyMarkers();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("HyperlinkWithMarker.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
        //Gets a list of company details
        private static List<Company> GetCompanyDetails()
        {
            List<Company> companyList = new List<Company>();

            Company company = new Company();
            company.Name = "Syncfusion";
            Hyperlink link = new Hyperlink("https://www.syncfusion.com", "", "", "Syncfusion", ExcelHyperLinkType.Url, null);
            company.Link = link;
            companyList.Add(company);

            company = new Company();
            company.Name = "Microsoft";
            link = new Hyperlink("https://www.microsoft.com", "", "", "Microsoft", ExcelHyperLinkType.Url, null);
            company.Link = link;
            companyList.Add(company);

            company = new Company();
            company.Name = "Google";
            link = new Hyperlink("https://www.google.com", "", "", "Google", ExcelHyperLinkType.Url, null);
            company.Link = link;
            companyList.Add(company);

            return companyList;
        }
        public class Hyperlink : IHyperLink
        {
            public IApplication Application { get; }
            public object Parent { get; }
            public string Address { get; set; }
            public string Name { get; }
            public IRange Range { get; }
            public string ScreenTip { get; set; }
            public string SubAddress { get; set; }
            public string TextToDisplay { get; set; }
            public ExcelHyperLinkType Type { get; set; }
            public IShape Shape { get; }
            public ExcelHyperlinkAttachedType AttachedType { get; }
            public byte[] Image { get; set; }

            public Hyperlink(string address, string subAddress, string screenTip, string textToDisplay, ExcelHyperLinkType type, byte[] image)
            {
                Address = address;
                ScreenTip = screenTip;
                SubAddress = subAddress;
                TextToDisplay = textToDisplay;
                Type = type;
                Image = image;
            }
        }

        public class Company
        {
            public string Name { get; set; }
            public Hyperlink Link { get; set; }
        }
    }
}

