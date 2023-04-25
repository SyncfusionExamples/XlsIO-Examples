using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;
using System;
using System.IO;

namespace LoanEMISchedule
{
    /// <summary>
    /// Represents a class that creates a loan EMI schedule in an Excel document using C#.
    /// </summary>
    class EMISchedule
    {
        // Fields to store the loan details
        private string m_BankName;
        private string m_CustomerName;
        private string m_AccountNumber;
        private string m_Tenure;
        private string m_InterestRate;
        private string m_LoanAmount;
        private string m_BorrowedDate;

        /// <summary>
        /// The entry point of the program.
        /// </summary>
        /// <param name="args">An array of command-line arguments.</param>
        public static void Main(string[] args)
        {
            // Register Syncfusion license
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("");

            // Create EMISchedule instance and generate loan schedule
            EMISchedule emiSchedule = new EMISchedule();
            emiSchedule.GenerateLoanSchedule();
        }

        /// <summary>
        /// Generates the loan schedule Excel document
        /// </summary>
        private void GenerateLoanSchedule()
        {
            // Initialize Excel Engine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                // Create new workbook and worksheet
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                // Get loan details from user
                GetLoanDetails();

                // Fill loan details in worksheet
                FillLoanDetails(sheet);

                // Add names for cells for easy referencing
                AddNamesForCells(workbook, sheet);

                // Calculate EMIs
                EMICalculation(sheet);

                // Apply cell styles
                ApplyCellStyles(sheet);

                // Enable sheet calculations
                sheet.EnableSheetCalculations();

                // Display EMI amount
                Console.WriteLine("Your EMI amount is.." + sheet["F10"].DisplayText);

                // Save workbook and close stream
                FileStream generatedExcel = new FileStream("../../../GeneratedOutput/Loan EMI Schedule.xlsx", FileMode.Create, FileAccess.Write);
                workbook.Version = ExcelVersion.Xlsx;
                workbook.SaveAs(generatedExcel);
                generatedExcel.Close();

                Console.WriteLine("Excel document generated successfully..");
            }
        }

        /// <summary>
        /// Gets the loan details from the user.
        /// </summary>
        private void GetLoanDetails()
        {
            Console.WriteLine("Enter the Bank name..");
            m_BankName = Console.ReadLine();
            Console.WriteLine("Enter the Customer Name..");
            m_CustomerName = Console.ReadLine();
            Console.WriteLine("Enter the Account Number..");
            m_AccountNumber = Console.ReadLine();
            Console.WriteLine("Enter the Tenure in months..");
            m_Tenure = Console.ReadLine();
            Console.WriteLine("Enter the Interest Rate per annum..");
            m_InterestRate = Console.ReadLine() + "%";
            Console.WriteLine("Enter the Loan Amount..");
            m_LoanAmount = Console.ReadLine();
            Console.WriteLine("Enter the Borrowed Date in the format MM-dd-yyyy..");
            m_BorrowedDate = Console.ReadLine();
            string[] dateValue = m_BorrowedDate.Split('-');
            DateTime formatDate = new DateTime(int.Parse(dateValue[2]), int.Parse(dateValue[0]), int.Parse(dateValue[1]));
            m_BorrowedDate = formatDate.ToString(System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat);
        }

        /// <summary>
        /// Fills the loan details in the worksheet.
        /// </summary>
        /// <param name="sheet">The worksheet to fill with the loan details.</param>
        private void FillLoanDetails(IWorksheet worksheet)
        {
            worksheet["A1"].Value = m_BankName;

            worksheet["A4"].Value = "Loan EMI Schedule";

            worksheet["A6"].Value = "Customer Name";
            worksheet["A8"].Value = "Account Number";
            worksheet["A10"].Value = "Tenure in months";
            worksheet["A12"].Value = "Interest";

            worksheet["B6"].Value = m_CustomerName;
            worksheet["B8"].Value = m_AccountNumber;
            worksheet["B10"].Value = m_Tenure;
            worksheet["B12"].Value = m_InterestRate;

            worksheet["E6"].Value = "Loan Amount";
            worksheet["E8"].Value = "Frequency";
            worksheet["E10"].Value = "EMI Amount";
            worksheet["E12"].Value = "Borrowed Date";

            worksheet["F6"].Value = m_LoanAmount;
            worksheet["F8"].Value = "Monthly";

            worksheet["F12"].Value = m_BorrowedDate; 

            worksheet["A15"].Value = "Payment No.";
            worksheet["B15"].Value = "Date";
            worksheet["C15"].Value = "Payment";
            worksheet["D15"].Value = "Principle";
            worksheet["E15"].Value = "Interest";
            worksheet["F15"].Value = "Outstanding Principle";
        }

        private void AddNamesForCells(IWorkbook workbook, IWorksheet sheet)
        {
            workbook.Names.Add("Interest", sheet["B12"]);
            workbook.Names.Add("Tenure", sheet["B10"]);
            workbook.Names.Add("LoanAmount", sheet["F6"]);
            workbook.Names.Add("BorrowedDate", sheet["F12"]);
        }
        /// <summary>
        /// Calculate EMI and generate EMI schedule
        /// </summary>
        /// <param name="sheet">Worksheet</param>
        private void EMICalculation(IWorksheet sheet)
        {
            sheet["F10"].Value = "=-PMT(Interest/12,Tenure, LoanAmount)";

            int totalEMIs = int.Parse(sheet["B10"].Value);

            sheet["A16"].Value = "1";
            sheet["B16"].Value = "=EDATE(BorrowedDate,1)";
            sheet["C16"].Value = sheet["F10"].Value;
            sheet["D16"].Value = "=$C16-$E16";
            sheet["E16"].Value = "=(Interest/12 * LoanAmount)";
            sheet["F16"].Value = "=LoanAmount-D16";

            sheet["A17"].Value = "=$A16 + 1";
            sheet["B17"].Value = "=EDATE($B16, 1)";
            sheet["C17"].Value = sheet["F10"].Value;
            sheet["D17"].Value = "=$C17-$E17";
            sheet["E17"].Value = "=(Interest/12 * $F16)";
            sheet["F17"].Value = "=$F16-$D17";

            IRange source = sheet["A17:F17"];
            IRange destination = sheet["A18:F18"];

            int count = totalEMIs - 2;
            while (count > 0)
            {
                source.CopyTo(destination);
                source = sheet[source.Row + 1, source.Column, source.LastRow + 1, source.LastColumn];
                destination = sheet[destination.Row + 1, destination.Column, destination.LastRow + 1, destination.LastColumn];
                count--;
            }

            IRange used = sheet.UsedRange;

            sheet[used.LastRow + 2, 4, used.LastRow + 2, 5].Merge();
            sheet[used.LastRow + 2, 4, used.LastRow + 2, 5].CellStyle.Font.Bold = true;
            sheet[used.LastRow + 2, 4, used.LastRow + 2, 5].Value = "Principle";
            sheet[used.LastRow + 3, 4, used.LastRow + 3, 5].Merge();
            sheet[used.LastRow + 3, 4, used.LastRow + 3, 5].CellStyle.Font.Bold = true;
            sheet[used.LastRow + 3, 4, used.LastRow + 3, 5].Value = "Interest";
            sheet[used.LastRow + 4, 4, used.LastRow + 4, 5].Merge();
            sheet[used.LastRow + 4, 4, used.LastRow + 4, 5].CellStyle.Font.Bold = true;
            sheet[used.LastRow + 4, 4, used.LastRow + 4, 5].Value = "Total Amount";

            sheet[used.LastRow + 2, 6, used.LastRow + 2, 6].Value = "=SUM(D16:D" + (16 + totalEMIs - 1) + ")";
            sheet[used.LastRow + 2, 6, used.LastRow + 2, 6].NumberFormat = "$#,###.00";
            sheet[used.LastRow + 3, 6, used.LastRow + 3, 6].Value = "=SUM(E16:E" + (16 + totalEMIs - 1) + ")";
            sheet[used.LastRow + 3, 6, used.LastRow + 3, 6].NumberFormat = "$#,###.00";
            sheet[used.LastRow + 4, 6, used.LastRow + 4, 6].Value = "=SUM(C16:C" + (16 + totalEMIs - 1) + ")";
            sheet[used.LastRow + 4, 6, used.LastRow + 4, 6].NumberFormat = "$#,###.00";

        }
        /// <summary>
        /// Apply cell styles
        /// </summary>
        /// <param name="sheet">Worksheet</param>
        private void ApplyCellStyles(IWorksheet sheet)
        {
            sheet.IsGridLinesVisible = false;

            sheet["A1:F2"].Merge();
            IStyle mergeArea = sheet["A1"].MergeArea.CellStyle;
            mergeArea.Font.Size = 18;
            mergeArea.Font.Bold = true;
            mergeArea.Font.Underline = ExcelUnderline.Single;

            sheet["A4:F4"].Merge();
            sheet["A4"].Value = "Loan EMI Schedule";
            mergeArea = sheet["A4"].MergeArea.CellStyle;
            mergeArea.Font.Size = 16;
            mergeArea.Font.Bold = true;

            sheet["A6:A12"].CellStyle.Font.Bold = true;
            sheet["E6:E12"].CellStyle.Font.Bold = true;
            sheet["F6"].NumberFormat = "$#,###.00";
            sheet["F10"].NumberFormat = "$#,###.00";

            sheet["A15:F15"].CellStyle.Font.Bold = true;
            sheet["A15:F15"].WrapText = true;
            sheet["A15:F15"].RowHeight = 31;           

            sheet.UsedRange.ColumnWidth = 15.5;

            IRange used = sheet.UsedRange;

            sheet[16,2, used.LastRow - 4, 2].NumberFormat = System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern;
            sheet[16, 3, used.LastRow - 4, 6].NumberFormat = "$#,###.00";

            sheet[15, 1, 15, 6].BorderAround(ExcelLineStyle.Thin);
            sheet[15, 1, 15, 6].BorderInside();
            sheet[16, 1, used.LastRow - 4, 6].BorderAround(ExcelLineStyle.Thin);
            sheet[16, 1, used.LastRow - 4, 6].Borders[ExcelBordersIndex.InsideVertical].LineStyle = ExcelLineStyle.Thin;

            sheet[6, 1, used.LastRow, 6].CellStyle.Font.Size = 12;

            sheet[16, 1, used.LastRow, 6].RowHeight = 24;
            sheet[1, 1, used.LastRow, 6].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
            sheet[1, 1, used.LastRow, 6].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
        }
    }
}
