using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.IO;

namespace Import_Nested_Collection
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
                marker.AddVariable("Customer", GetSalesReports());

                //Process the markers in the template
                marker.ApplyMarkers();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ImportNestedCollection.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
        //Get sales Report class
        public static List<Customer> GetSalesReports()
        {
            List<Customer> reports = new List<Customer>();

            List<Order> orders = new List<Order>();
            orders.Add(new Order(1408, 451.75));
            orders.Add(new Order(1278, 340.00));
            orders.Add(new Order(1123, 290.50));

            Customer c1 = new Customer(002107, "Andy Bernard", 45);
            c1.Orders = orders;
            Customer c2 = new Customer(011564, "Jim Halpert", 34);
            c2.Orders = orders;
            Customer c3 = new Customer(002097, "Karen Fillippelli", 35);
            c3.Orders = orders;
            Customer c4 = new Customer(001846, "Phyllis Lapin", 37);
            c4.Orders = orders;
            Customer c5 = new Customer(012167, "Stanley Hudson", 41);
            c5.Orders = orders;

            reports.Add(c1);
            reports.Add(c2);
            reports.Add(c3);
            reports.Add(c4);
            reports.Add(c5);

            return reports;
        }

        //Customer details
        public partial class Customer
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public int Age { get; set; }
            public IList<Order> Orders { get; set; }
            public Customer(int id, string name, int age)
            {
                Id = id;
                Name = name;
                Age = age;
            }
        }
        //Order details
        public partial class Order
        {
            public int Order_Id { get; set; }
            public double Price { get; set; }

            public Order(int id, double price)
            {
                Order_Id = id;
                Price = price;
            }
        }
    }
}






