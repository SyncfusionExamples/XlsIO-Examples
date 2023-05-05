using Syncfusion.XlsIO;
using System;
using System.IO;
using System.Threading;

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
        private long m_AccountNumber;
        private int m_Tenure;
        private double m_InterestRate;
        private long m_LoanAmount;
        private DateTime m_BorrowedDate;

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
            emiSchedule.GenerateLoanEMISchedule();
        }

        /// <summary>
        /// Generates the loan schedule Excel document
        /// </summary>
        private void GenerateLoanEMISchedule()
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

                // Calculate EMI
                CalculateEMI(sheet, m_BankName, m_AccountNumber, m_CustomerName, m_InterestRate, m_LoanAmount, m_Tenure, m_BorrowedDate);

                // Display the EMI amount
                Console.WriteLine("Your EMI amount is.." + sheet["F10"].DisplayText);

                // Save workbook and close stream
                Directory.CreateDirectory("../../../GeneratedOutput");
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
            try
            {
                Console.WriteLine("Enter the Bank name..");
                m_BankName = Console.ReadLine();

                Console.WriteLine("Enter the Customer Name..");
                m_CustomerName = Console.ReadLine();

                Console.WriteLine("Enter the Account Number..");
                m_AccountNumber = long.Parse(Console.ReadLine());

                Console.WriteLine("Enter the Tenure in months..");
                m_Tenure = int.Parse(Console.ReadLine());

                Console.WriteLine("Enter the Interest Rate per annum..");
                m_InterestRate = double.Parse(Console.ReadLine());

                Console.WriteLine("Enter the Loan Amount..");
                m_LoanAmount = long.Parse(Console.ReadLine());

                Console.WriteLine("Enter the Borrowed Date in the format MM-dd-yyyy..");
                string date = Console.ReadLine();
                string[] dateValue = date.Split('-');
                m_BorrowedDate = new DateTime(int.Parse(dateValue[2]), int.Parse(dateValue[0]), int.Parse(dateValue[1]));

            }
            catch (Exception ex)
            {
                Console.WriteLine("Please enter valid input  " + ex.ToString());
                GetLoanDetails();
            }

        }
        /// <summary>
        /// Calculate EMI and generate EMI schedule
        /// </summary>
        /// <param name="sheet">Worksheet</param>
        private static void CalculateEMI(IWorksheet sheet, string bankName, long accountNumber, string customerName, double interestRate, long loanAmount, int tenureInMonths, DateTime borrowedDate)
        {
            sheet["A1"].Value = bankName;

            sheet["A4"].Value = "Loan EMI Schedule";

            sheet["A6"].Value = "Customer Name";
            sheet["A8"].Value = "Account Number";
            sheet["A10"].Value = "Tenure in months";
            sheet["A12"].Value = "Interest";

            sheet["B6"].Text = customerName;
            sheet["B8"].Number = accountNumber;
            sheet["B10"].Number = tenureInMonths;
            sheet["B12"].Number = interestRate/100;

            sheet["E6"].Value = "Loan Amount";
            sheet["E8"].Value = "Frequency";
            sheet["E10"].Value = "EMI Amount";
            sheet["E12"].Value = "Borrowed Date";

            sheet["F6"].Number = loanAmount;
            sheet["F8"].Value = "Monthly";
            sheet["F12"].DateTime = borrowedDate;

            sheet["A15"].Value = "Payment No.";
            sheet["B15"].Value = "Date";
            sheet["C15"].Value = "Payment";
            sheet["D15"].Value = "Principle";
            sheet["E15"].Value = "Interest";
            sheet["F15"].Value = "Outstanding Principle";

            sheet.Workbook.Names.Add("Interest", sheet["B12"]);
            sheet.Workbook.Names.Add("Tenure", sheet["B10"]);
            sheet.Workbook.Names.Add("LoanAmount", sheet["F6"]);
            sheet.Workbook.Names.Add("BorrowedDate", sheet["F12"]);

            sheet["F10"].Formula = "=-PMT(Interest/12,Tenure, LoanAmount)";

            sheet.EnableSheetCalculations();

            double emi = double.Parse(sheet["F10"].CalculatedValue.ToString());

            double balance = loanAmount;

            double totalInterestPaid = 0;

            for (int i = 1; i <= tenureInMonths; i++)
            {
                double interest = balance * (interestRate/100)/12;
                double principal = emi - interest;
                balance -= principal;

                totalInterestPaid += interest;

                sheet[15 + i, 1].Number = i;
                sheet[15 + i, 2].Formula = "=EDATE(BorrowedDate," + i + ")";
                sheet[15 + i, 3].Number = emi;
                sheet[15 + i, 4].Number = principal;
                sheet[15 + i, 5].Number = interest;
                sheet[15 + i, 6].Number = balance;
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

            sheet[used.LastRow + 2, 6, used.LastRow + 2, 6].Number = loanAmount;
            sheet[used.LastRow + 2, 6, used.LastRow + 2, 6].NumberFormat = "$#,###.00";
            sheet[used.LastRow + 3, 6, used.LastRow + 3, 6].Number = totalInterestPaid;
            sheet[used.LastRow + 3, 6, used.LastRow + 3, 6].NumberFormat = "$#,###.00";
            sheet[used.LastRow + 4, 6, used.LastRow + 4, 6].Number = totalInterestPaid + loanAmount;
            sheet[used.LastRow + 4, 6, used.LastRow + 4, 6].NumberFormat = "$#,###.00";

            //Apply styles to the cells
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

            sheet["B12"].NumberFormat = "0.0%";
            sheet["F12"].NumberFormat = Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern;

            sheet["A15:F15"].CellStyle.Font.Bold = true;
            sheet["A15:F15"].WrapText = true;
            sheet["A15:F15"].RowHeight = 31;           

            sheet.UsedRange.ColumnWidth = 15.5;

            used = sheet.UsedRange;

            sheet[16,2, used.LastRow - 4, 2].NumberFormat = Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern;
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
