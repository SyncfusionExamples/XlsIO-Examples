﻿using System.Collections.Generic;
using System.IO;
using Syncfusion.XlsIO;

namespace Worksheet_to_Nested_Class
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Map column headers in worksheet with class properties. 
                Dictionary<string, string> mappingProperties = new Dictionary<string, string>();
                mappingProperties.Add("Customer ID", "CustId");
                mappingProperties.Add("Customer Name", "CustName");
                mappingProperties.Add("Customer Age", "CustAge");
                mappingProperties.Add("Order ID", "CustOrder.Order_Id");
                mappingProperties.Add("Order Price", "CustOrder.Price");

                //Export worksheet data into nested class Objects.
                List<Customer> nestedClassObjects = worksheet.ExportData<Customer>(1, 1, 10, 5, mappingProperties);
            }
        }
    }
    //Customer details class
    public partial class Customer
    {
        public int CustId { get; set; }
        public string CustName { get; set; }
        public int CustAge { get; set; }
        public Order CustOrder { get; set; }
        public Customer()
        {

        }
    }

    //Order details class
    public partial class Order
    {
        public string Order_Id { get; set; }
        public double Price { get; set; }
        public Order()
        {

        }
    }
}





