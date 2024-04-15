using System.Data;
using Syncfusion.XlsIO;
using Newtonsoft.Json;


namespace ExportDataTableWithBase64Images
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                string mergeData = GetData();
                DataTable dtExport = (DataTable)JsonConvert.DeserializeObject(mergeData, typeof(DataTable));

                //Change the image column datatype to byte[]
                ChangeColumnType(dtExport, "Signature", typeof(byte[]));

                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Create a new workbook
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx",FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Create template marker processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                //Add marker variable
                marker.AddVariable("GeneratedData", dtExport);

                //Apply markers
                marker.ApplyMarkers();

                //Autofit the columns
                worksheet.UsedRange.AutofitColumns();

                //Save the workbook
                FileStream fileStream = new FileStream("Output.xlsx",FileMode.Create,FileAccess.Write);
                workbook.SaveAs(fileStream);

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Output.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }

        /// <summary>
        /// Change the image column type to byte[]
        /// </summary>
        /// <param name="dataTable">data table</param>
        /// <param name="columnName">column name of the datatable</param>
        /// <param name="newType">new type of the column</param>
        public static void ChangeColumnType(DataTable dataTable, string columnName, Type newType)
        {
            if (dataTable.Columns.Contains(columnName))
            {
                // Create a new column with the new data type
                DataColumn newColumn = new DataColumn("Images", newType);

                // Add the new column to the DataTable
                dataTable.Columns.Add(newColumn);

                // Copy data from the old column to the new column
                foreach (DataRow row in dataTable.Rows)
                {
                    // convert the base64 image to byte[]
                    byte[] imageData = Convert.FromBase64String(row[columnName].ToString());
                    row[columnName] = imageData;

                    // Set the image data for the new column
                    row[newColumn] = imageData;
                }

                // Remove the old column
                dataTable.Columns.Remove(columnName);
                newColumn.ColumnName = columnName;
            }
        }

        /// <summary>
        /// Json Data for populating a datatable
        /// </summary>
        /// <returns></returns>
        private static string GetData()
        {
            return "[{\"Name\": \"John\",\"LastName\": \"Doe\",\"Address\": \"123 Main St\",\"City\": \"New York\",\"Telephone\": \"987-654-3210\",\"Email\": \"john.doe@example.com\",\"Signature\": \"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\"}]";
        }
    }
}