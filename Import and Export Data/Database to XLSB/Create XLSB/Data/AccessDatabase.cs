using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Web;

namespace ExcelDataExport.Data
{
    /// <summary>
    /// Creates App_Data\SampleDB.mdb (Microsoft Access) with a Products table on first use
    /// and reads it into a DataTable.
    /// </summary>
    public static class AccessDatabase
    {
        private static readonly object SyncLock = new object();

        // Provider used when the .mdb file has to be created (keep in sync with Web.config).
        private static string ConnectionString
        {
            get { return ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString; }
        }

        private static string DatabasePath
        {
            get { return HttpContext.Current.Server.MapPath("~/App_Data/SampleDB.mdb"); }
        }

        private static string ResolvedConnectionString
        {
            get { return ConnectionString.Replace("|DataDirectory|", HttpContext.Current.Server.MapPath("~/App_Data")); }
        }

        public static DataTable GetProducts()
        {
            EnsureDatabase();

            const string query =
                "SELECT ProductID, ProductName, Category, UnitPrice, UnitsInStock, DateAdded " +
                "FROM Products ORDER BY ProductID";

            DataTable table = new DataTable("Products");
            using (OleDbConnection connection = new OleDbConnection(ResolvedConnectionString))
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
            {
                adapter.Fill(table);
            }
            return table;
        }

        private static void EnsureDatabase()
        {
            lock (SyncLock)
            {
                if (File.Exists(DatabasePath))
                {
                    return;
                }

                Directory.CreateDirectory(Path.GetDirectoryName(DatabasePath));

                // 1. Create an empty .mdb file with ADOX (late bound, no COM reference needed)
                Type catalogType = Type.GetTypeFromProgID("ADOX.Catalog");
                if (catalogType == null)
                {
                    throw new InvalidOperationException(
                        "ADOX (msadox.dll) is not available. Copy a SampleDB.mdb file into App_Data manually.");
                }

                dynamic catalog = Activator.CreateInstance(catalogType);
                try
                {
                    catalog.Create(ResolvedConnectionString);
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(catalog);
                }

                // 2. Create the table and insert sample rows
                using (OleDbConnection connection = new OleDbConnection(ResolvedConnectionString))
                {
                    connection.Open();

                    using (OleDbCommand create = new OleDbCommand(
                        "CREATE TABLE Products (" +
                        "ProductID COUNTER PRIMARY KEY, " +
                        "ProductName TEXT(100) NOT NULL, " +
                        "Category TEXT(50) NOT NULL, " +
                        "UnitPrice CURRENCY NOT NULL, " +
                        "UnitsInStock LONG NOT NULL, " +
                        "DateAdded DATETIME NOT NULL)", connection))
                    {
                        create.ExecuteNonQuery();
                    }

                    object[][] rows =
                    {
                        new object[] { "Laptop Computer",     "Electronics", 999.99m, 15 },
                        new object[] { "Wireless Mouse",      "Electronics",  29.99m, 150 },
                        new object[] { "USB-C Cable",         "Accessories",  12.99m, 500 },
                        new object[] { "Monitor 27inch",      "Electronics", 349.99m, 25 },
                        new object[] { "Keyboard Mechanical", "Electronics",  89.99m, 80 },
                        new object[] { "Desk Lamp LED",       "Accessories",  44.99m, 120 },
                        new object[] { "Office Chair",        "Furniture",   299.99m, 10 },
                        new object[] { "Standing Desk",       "Furniture",   599.99m, 5 },
                        new object[] { "USB Hub",             "Accessories",  39.99m, 200 },
                        new object[] { "Webcam HD",           "Electronics",  79.99m, 60 }
                    };

                    foreach (object[] row in rows)
                    {
                        using (OleDbCommand insert = new OleDbCommand(
                            "INSERT INTO Products (ProductName, Category, UnitPrice, UnitsInStock, DateAdded) " +
                            "VALUES (?, ?, ?, ?, ?)", connection))
                        {
                            insert.Parameters.Add("ProductName", OleDbType.VarWChar, 100).Value = row[0];
                            insert.Parameters.Add("Category", OleDbType.VarWChar, 50).Value = row[1];
                            insert.Parameters.Add("UnitPrice", OleDbType.Currency).Value = row[2];
                            insert.Parameters.Add("UnitsInStock", OleDbType.Integer).Value = row[3];
                            insert.Parameters.Add("DateAdded", OleDbType.Date).Value = DateTime.Now;
                            insert.ExecuteNonQuery();
                        }
                    }
                }
            }
        }
    }
}
