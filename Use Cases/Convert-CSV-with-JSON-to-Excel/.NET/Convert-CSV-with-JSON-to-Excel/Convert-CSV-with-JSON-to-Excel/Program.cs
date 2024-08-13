using Syncfusion.XlsIO;

namespace Convert_CSV_with_JSON_to_Excel
{
    class Program
    {
        public static void Main(string[] args)
        {
            bool startConcatenation = false;
            string startCellAddress = "";
            List<string> concatenatedValues = new List<string>();

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Loads an CSV file
                FileStream fileStream = new FileStream(@"../../../Data/InputTemplate.csv", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream, ";");
                IWorksheet worksheet = workbook.Worksheets[0];

                for (int row = 1; row <= worksheet.UsedRange.LastRow; row++)
                {
                    for (int col = 1; col <= worksheet.UsedRange.LastColumn; col++)
                    {
                        IRange cell = worksheet.UsedRange[row, col];
                        string cellValue = cell.DisplayText;

                        //Checks the cellvalue starts with {
                        if (cellValue.StartsWith("\"{\n") || cellValue.StartsWith(" \"{"))
                        {
                            startConcatenation = true;
                            startCellAddress = cell.AddressLocal;
                            concatenatedValues.Add(cellValue);
                            continue; //Skip to the next cell
                        }

                        //Concatenate the JSON value to the list
                        if (startConcatenation && !string.IsNullOrEmpty(cellValue))
                        {
                            
                            concatenatedValues.Add(cellValue);
                            cell.Clear();
                        }

                        //Update the JSON value in the respective cell
                        if (cellValue.Contains("}\""))
                        {
                            startConcatenation = false;
                            //Concatenate the values 
                            string concatenatedValue = string.Join(" ", concatenatedValues);

                            //Update the corresponding cell with the concatenated value
                            worksheet.Range[startCellAddress].Value = concatenatedValue;

                            //Clear the list for the next iteration
                            concatenatedValues.Clear();

                            //Reset the start cell address for the next iteration
                            startCellAddress = "";
                        }
                    }
                }

                //Check for blank row
                List<int> rowsToDelete = new List<int>();
                for (int row = 1; row <= worksheet.UsedRange.LastRow; row++)
                {
                    bool isRowEmpty = true;
                    for (int col = 1; col <= worksheet.UsedRange.LastColumn; col++)
                    {
                        IRange cell = worksheet.Range[row, col];
                        if (!string.IsNullOrEmpty(cell.Value))
                        {
                            isRowEmpty = false;
                            break;
                        }
                    }
                    if (isRowEmpty)
                    {
                        rowsToDelete.Add(row);
                    }
                }

                //Delete the blank rows
                for (int i = 0; i < rowsToDelete.Count; i++)
                {
                    worksheet.DeleteRow(rowsToDelete[i] - i);
                }

                worksheet.UsedRange.WrapText = false;
                worksheet.UsedRange.AutofitColumns();
                FileStream outputStream = new FileStream("output.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
            }
        }
    }
}