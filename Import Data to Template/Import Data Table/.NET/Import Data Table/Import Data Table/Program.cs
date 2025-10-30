﻿using System;
using System.Data;
using System.IO;
using Syncfusion.XlsIO;

namespace Import_Data_Table
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);                
                IWorksheet worksheet = workbook.Worksheets[0];

                //Create Template Marker Processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                //Create an instance for Data table
                DataTable reports = new DataTable();

                //Add value to data table
                reports.Columns.Add("SalesPerson");
                reports.Columns.Add("FromDate", typeof(DateTime));
                reports.Columns.Add("ToDate", typeof(DateTime));

                reports.Rows.Add("Andy Bernard", new DateTime(2014, 09, 08), new DateTime(2014, 09, 11));
                reports.Rows.Add("Jim Halpert", new DateTime(2014, 09, 11), new DateTime(2014, 09, 15));
                reports.Rows.Add("Karen Fillippelli", new DateTime(2014, 09, 15), new DateTime(2014, 09, 20));
                reports.Rows.Add("Phyllis Lapin", new DateTime(2014, 09, 21), new DateTime(2014, 09, 25));
                reports.Rows.Add("Stanley Hudson", new DateTime(2014, 09, 26), new DateTime(2014, 09, 30));

                //Add collection to marker variable
                marker.AddVariable("Reports", reports, VariableTypeAction.DetectNumberFormat);

                //Process the markers in the template
                marker.ApplyMarkers();

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ImportDataTable.xlsx"));
                #endregion
            }
        }
    }
}