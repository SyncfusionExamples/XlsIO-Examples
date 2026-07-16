using Syncfusion.XlsIO;
using Syncfusion.Drawing;
using System.IO;

namespace Form_Controls
{
    class Program
    {
        static void Main()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                // Create workbook with 2 worksheets
                IWorkbook workbook = application.Workbooks.Create(2);
                IWorksheet sheet = workbook.Worksheets[0];
                IWorksheet sheet2 = workbook.Worksheets[1];

                // Payment options for combo box
                string[] onlinePayments = { "Credit Card", "Net Banking" };
                for (int i = 0; i < onlinePayments.Length; i++)
                    sheet2.SetValue(i + 1, 1, onlinePayments[i]);

                // CONTACT SALES section
                sheet[2, 3].Text = "CONTACT SALES";
                sheet[2, 3].CellStyle.Font.Bold = true;
                sheet[2, 3].CellStyle.Font.Size = 14;
                sheet[2, 3].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;

                sheet[4, 3].Text = "Phone";
                sheet[4, 3].CellStyle.Font.Bold = true;
                sheet[5, 3].Text = "Toll Free";
                sheet[5, 5].Text = "1-888-9DOTNET";
                sheet[6, 5].Text = "1-888-936-8638";
                sheet[7, 5].Text = "1-919-481-1974";

                sheet[8, 3].Text = "Fax";
                sheet[8, 5].Text = "1-919-573-0306";

                sheet[9, 3].Text = "Email";
                sheet[10, 3].Text = "Sales";

                IHyperLink link = sheet.HyperLinks.Add(sheet[10, 5]);
                link.Type = ExcelHyperLinkType.Url;
                link.Address = "mailto:sales@syncfusion.com";
                sheet[10, 5].Text = "sales@syncfusion.com";
                sheet[10, 5].CellStyle.Font.Color = ExcelKnownColors.Blue;
                sheet[10, 5].CellStyle.Font.Underline = ExcelUnderline.Single;

                sheet[12, 3].Text = "Please fill out all required fields.";
                sheet[12, 3].CellStyle.Font.Italic = true;
                sheet[12, 3].CellStyle.Font.Color = ExcelKnownColors.Grey_80_percent;

                // Form controls
                sheet[14, 5].Text = "First Name*"; sheet[14, 5].CellStyle.Font.Bold = true;
                sheet[14, 8].Text = "Last Name*"; sheet[14, 8].CellStyle.Font.Bold = true;

                ITextBoxShape textBoxShape = sheet.TextBoxes.AddTextBox(15, 5, 23, 190);
                textBoxShape.Fill.FillType = ExcelFillType.SolidColor;
                textBoxShape.Fill.ForeColor = Color.LightYellow;

                textBoxShape = sheet.TextBoxes.AddTextBox(15, 8, 23, 195);
                textBoxShape.Fill.FillType = ExcelFillType.SolidColor;
                textBoxShape.Fill.ForeColor = Color.LightYellow;

                sheet[17, 3].Text = "Company*"; 
                textBoxShape = sheet.TextBoxes.AddTextBox(17, 5, 23, 385);
                textBoxShape.Fill.FillType = ExcelFillType.SolidColor;
                textBoxShape.Fill.ForeColor = Color.LightYellow;

                sheet[19, 3].Text = "Phone*"; 
                textBoxShape = sheet.TextBoxes.AddTextBox(19, 5, 23, 385);
                textBoxShape.Fill.FillType = ExcelFillType.SolidColor;
                textBoxShape.Fill.ForeColor = Color.LightYellow;

                sheet[21, 3].Text = "Email*";
                textBoxShape = sheet.TextBoxes.AddTextBox(21, 5, 23, 385);
                textBoxShape.Fill.FillType = ExcelFillType.SolidColor;
                textBoxShape.Fill.ForeColor = Color.LightYellow;

                sheet[23, 3].Text = "Website";
                textBoxShape = sheet.TextBoxes.AddTextBox(23, 5, 23, 385);
                textBoxShape.Fill.FillType = ExcelFillType.SolidColor;
                textBoxShape.Fill.ForeColor = Color.LightYellow;

                // Multiple products
                sheet[25, 3].Text = "Multiple products?";
                ICheckBoxShape chkBoxProducts = sheet.CheckBoxes.AddCheckBox(25, 5, 20, 75);
                chkBoxProducts.CheckState = ExcelCheckState.Mixed;

                // Product(s) section
                sheet[27, 3, 28, 3].Merge();
                sheet[27, 3].Text = "Product(s)*";
                sheet[27, 3].MergeArea.CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                ICheckBoxShape chkBoxProduct;
                chkBoxProduct = sheet.CheckBoxes.AddCheckBox(27, 5, 20, 75); chkBoxProduct.Text = "Studio";
                chkBoxProduct = sheet.CheckBoxes.AddCheckBox(27, 6, 20, 75); chkBoxProduct.Text = "Calculate";
                chkBoxProduct = sheet.CheckBoxes.AddCheckBox(27, 7, 20, 75); chkBoxProduct.Text = "Chart";
                chkBoxProduct = sheet.CheckBoxes.AddCheckBox(27, 8, 20, 75); chkBoxProduct.Text = "Diagram";
                chkBoxProduct = sheet.CheckBoxes.AddCheckBox(27, 9, 20, 75); chkBoxProduct.Text = "Edit";
                chkBoxProduct = sheet.CheckBoxes.AddCheckBox(27, 10, 20, 75); chkBoxProduct.Text = "XlsIO";
                chkBoxProduct = sheet.CheckBoxes.AddCheckBox(28, 5, 20, 75); chkBoxProduct.Text = "Grid";
                chkBoxProduct = sheet.CheckBoxes.AddCheckBox(28, 6, 20, 75); chkBoxProduct.Text = "Grouping";
                chkBoxProduct = sheet.CheckBoxes.AddCheckBox(28, 7, 20, 75); chkBoxProduct.Text = "HTMLUI";
                chkBoxProduct = sheet.CheckBoxes.AddCheckBox(28, 8, 20, 75); chkBoxProduct.Text = "PDF";
                chkBoxProduct = sheet.CheckBoxes.AddCheckBox(28, 9, 20, 75); chkBoxProduct.Text = "Tools";
                chkBoxProduct = sheet.CheckBoxes.AddCheckBox(28, 10, 20, 75); chkBoxProduct.Text = "DocIO";

                // Link checkboxes to hidden cells and formulas
                GenerateFormula(excelEngine);

                // Selected products count
                sheet[30, 3].Text = "Selected Products Count";
                sheet[30, 5].Formula = "SUM(AA2:AA13)";
                sheet[30, 5].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                // Additional Information
                sheet[35, 3].Text = "Additional Information";
                sheet.TextBoxes.AddTextBox(32, 5, 150, 385);

                // Combo box
                sheet[43, 3].Text = "Online Payment";
                IComboBoxShape comboBox1 = sheet.ComboBoxes.AddComboBox(43, 5, 20, 100);
                comboBox1.ListFillRange = sheet2["A1:A2"];
                comboBox1.SelectedIndex = 1;

                // Option buttons
                sheet[46, 3].Text = "Card Type";
                IOptionButtonShape optionButton1 = sheet.OptionButtons.AddOptionButton(46, 5);
                optionButton1.Text = "American Express";
                optionButton1.CheckState = ExcelCheckState.Checked;

                optionButton1 = sheet.OptionButtons.AddOptionButton(46, 7);
                optionButton1.Text = "Master Card";

                optionButton1 = sheet.OptionButtons.AddOptionButton(46, 9);
                optionButton1.Text = "Visa";

                // Styling
                sheet.Columns[0].AutofitColumns();
                sheet.Columns[3].ColumnWidth = 12;
                sheet.Columns[4].ColumnWidth = 10;
                sheet.Columns[5].ColumnWidth = 10;
                sheet.IsGridLinesVisible = false;

                // Delete unused rows
                sheet.DeleteRow(40);
                sheet.DeleteRow(41);
                sheet.DeleteRow(42);
                sheet.DeleteRow(45);

                // Save workbook
                workbook.SaveAs(Path.GetFullPath("Output/FormControls.xlsx"));
            }
        }

        private static void GenerateFormula(ExcelEngine excelEngine)
        {
            IWorksheet worksheet = excelEngine.Excel.Workbooks[0].Worksheets[0];
            ICheckBoxes checkBoxes = worksheet.CheckBoxes;
            string formula;

            for (int i = 1; i < checkBoxes.Count; i++)
            {
                IRange range = worksheet["Z" + (i + 1)];
                checkBoxes[i].LinkedCell = range;
                formula = "IF(" + range.AddressLocal + ",1,0)";
                worksheet["AA" + (i + 1)].Formula = formula;
            }
        }
    }
}