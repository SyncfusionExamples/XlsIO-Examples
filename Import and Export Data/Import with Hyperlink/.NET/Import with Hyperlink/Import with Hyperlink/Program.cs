using System.IO;
using Syncfusion.XlsIO;
using System.Collections.Generic;

namespace Import_with_Hyperlink
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Import the data to worksheet
                IList<Company> reports = GetCompanyDetails();
                worksheet.ImportData(reports, 2, 1, false);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ImportData.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ImportData.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
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
