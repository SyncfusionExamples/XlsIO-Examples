using Syncfusion.Drawing;
using Syncfusion.XlsIO;
using System;

public class SupportTicketFormWithControls
{
    public static void Main()
    {
        //Create Form for Support Tickets
        CreateExcelFormWithControls("SupportTicketForm.xlsx");

        //Use below method to read filled Excel Form
        //ReadExcelForms("SupportTicketForm.xlsx");
    }

    /// <summary>
    /// Read Excel Form
    /// </summary>
    /// <param name="form">Excel File Path</param>
    private static void ReadExcelForms(string form)
    {
        //Initialize Excel Engine
        using (ExcelEngine engine = new ExcelEngine())
        {
            //Access Excel application
            IApplication application = engine.Excel;

            //Set the default Excel version
            application.DefaultVersion = ExcelVersion.Xlsx;

            //Create Excel workbook
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet responseSheet = workbook.Worksheets[0];

            //Set sheet name
            responseSheet.Name = "Responses";

            //Create response table
            BuildResponsesTable(responseSheet);

            //Read Excel form
            ReadResponses(form, responseSheet);

            //Save the workbook
            workbook.SaveAs("SupportTicketResponses.xlsx");
        }
    }

    /// <summary>
    /// Read responses from the form and populate in the response sheet
    /// </summary>
    /// <param name="form">Excel form</param>
    /// <param name="responseSheet">Response worksheet</param>
    private static void ReadResponses(string form, IWorksheet responseSheet)
    {
        // Initialize Excel Engine
        using (ExcelEngine engine = new ExcelEngine())
        {
            // Access Excel application
            IApplication application = engine.Excel;

            // Set the default Excel version
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Open the Excel form
            IWorkbook formWb = application.Workbooks.Open(form);
            IWorksheet formSheet = formWb.Worksheets[0];

            // Access last row of the table
            int lastRow = responseSheet.ListObjects[0].Location.LastRow;
            responseSheet.InsertRow(lastRow + 1);
            int newRow = lastRow;

            // Set ticket id
            responseSheet.Range[$"A{newRow}"].Number = newRow - 1;

            // Set Response created date
            responseSheet.Range[$"B{newRow}"].DateTime = DateTime.Now;

            // Set Ticket created date
            responseSheet.Range[$"C{newRow}"].DateTime = formSheet.Range["D5"].DateTime;

            // Set customer name
            responseSheet.Range[$"D{newRow}"].Text = formSheet.TextBoxes["CustomerNameTextBox"].Text;

            // Set Email address
            responseSheet.Range[$"E{newRow}"].Text = formSheet.TextBoxes["EmailTextBox"].Text;

            // Set Phone number
            responseSheet.Range[$"F{newRow}"].Text = formSheet.TextBoxes["PhoneTextBox"].Text;

            // Set communicated channel
            responseSheet.Range[$"G{newRow}"].Text = formSheet.ComboBoxes[0].SelectedValue;

            // Set Category
            responseSheet.Range[$"H{newRow}"].Text = formSheet.ComboBoxes[1].SelectedValue;

            // Set Priority
            responseSheet.Range[$"I{newRow}"].Text = formSheet.OptionButtons["Priority"].LinkedCell.Text;

            // Set whether follow up needed.
            responseSheet.Range[$"J{newRow}"].Text = formSheet.CheckBoxes["FollowUpCheckBox"].CheckState == ExcelCheckState.Checked ? "Needed" : "No Need"; // RequiresFollowUp

            // Set ticket summary
            responseSheet.Range[$"K{newRow}"].Text = formSheet.TextBoxes["IssueSummaryTextBox"].Text;

            // Set details
            responseSheet.Range[$"L{newRow}"].Text = formSheet.TextBoxes["DetailsTextBox"].Text;

            responseSheet.UsedRange.AutofitColumns();
        }
    }

    /// <summary>
    /// Build response table to store the form responses
    /// </summary>
    /// <param name="worksheet">Worksheet</param>
    private static void BuildResponsesTable(IWorksheet worksheet)
    {
        // Set the required form categories
        worksheet.Range["A1"].Text = "Ticket ID";
        worksheet.Range["B1"].Text = "CreatedOn";
        worksheet.Range["C1"].Text = "TicketDate";
        worksheet.Range["D1"].Text = "CustomerName";
        worksheet.Range["E1"].Text = "Email";
        worksheet.Range["F1"].Text = "Phone";
        worksheet.Range["G1"].Text = "Channel";
        worksheet.Range["H1"].Text = "Category";
        worksheet.Range["I1"].Text = "Priority";
        worksheet.Range["J1"].Text = "RequiresFollowUp";
        worksheet.Range["K1"].Text = "IssueSummary";
        worksheet.Range["L1"].Text = "Details";

        // Create a table
        var table = worksheet.ListObjects.Create("ResponsesTable", worksheet.Range["A1:L1"]);
        table.BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium9;
        worksheet.UsedRange.AutofitColumns();
    }

    /// <summary>
    ///  Create Excel Form with various controls
    /// </summary>
    /// <param name="filePath">Excel file path</param>
    private static void CreateExcelFormWithControls(string filePath)
    {
        // Initialize Excel Engine
        using (ExcelEngine engine = new ExcelEngine())
        {
            // Access Excel application
            IApplication application = engine.Excel;

            // Set the default Excel version
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Create workbook with two worksheets
            IWorkbook workbook = application.Workbooks.Create(2);
            IWorksheet formSheet = workbook.Worksheets[0];

            // Set worksheet name
            formSheet.Name = "Form";

            // Access lookup worksheet and set name
            IWorksheet lookupSheet = workbook.Worksheets[1];
            lookupSheet.Name = "Lookups";

            // Create lookup references for form
            BuildLookupsAndNames(workbook, lookupSheet);

            // Create Excel form
            BuildFormLayout(formSheet);

            // Save the Form as Excel workbook
            workbook.SaveAs("SupportTicketForm.xlsx");
        }
    }

    /// <summary>
    /// Build Look up references and named ranges
    /// </summary>
    /// <param name="workbook">Workbook</param>
    /// <param name="lookupSheet">Lookup Worksheet</param>
    private static void BuildLookupsAndNames(IWorkbook workbook, IWorksheet lookupSheet)
    {
        // Set Titles
        lookupSheet.Range["A1"].Text = "Channels";
        lookupSheet.Range["B1"].Text = "Categories";
        lookupSheet.Range["C1"].Text = "Priorities";

        // Add Channels
        lookupSheet.Range["A2"].Text = "Email";
        lookupSheet.Range["A3"].Text = "Phone";
        lookupSheet.Range["A4"].Text = "Chat";
        lookupSheet.Range["A5"].Text = "Web";

        // Add Categories
        lookupSheet.Range["B2"].Text = "Account";
        lookupSheet.Range["B3"].Text = "Billing";
        lookupSheet.Range["B4"].Text = "Technical";
        lookupSheet.Range["B5"].Text = "Other";


        // Add Priorities
        lookupSheet.Range["C2"].Text = "Low";
        lookupSheet.Range["C3"].Text = "Medium";
        lookupSheet.Range["C4"].Text = "High";
        lookupSheet.Range["C5"].Text = "Critical";

        // Create named range references
        workbook.Names.Add("Channels").RefersToRange = lookupSheet.Range["A2:A5"];
        workbook.Names.Add("CategoryList").RefersToRange = lookupSheet.Range["B2:B5"];
        workbook.Names.Add("Priorities").RefersToRange = lookupSheet.Range["C2:C5"];

        // Autofit the columns
        lookupSheet.UsedRange.AutofitColumns();

        // Protect the Excel worksheet
        lookupSheet.Protect("ProtectLookUp");
    }

    /// <summary>
    /// Layout the form and add form controls
    /// </summary>
    /// <param name="formSheet">Form worksheet</param>
    private static void BuildFormLayout(IWorksheet formSheet)
    {
        // Create Form Titles    
        formSheet.Range["B2"].Text = "Support Ticket Intake";
        formSheet.Range["B2"].CellStyle.Font.Bold = true;
        formSheet.Range["B2"].CellStyle.Font.Size = 18;
        formSheet.Range["B2"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
        formSheet.Range["B2:F2"].Merge();

        formSheet.Range["B3:F3"].Merge();
        formSheet.Range["B3"].Text = "Fill the form using the controls. Required fields marked *.";
        formSheet.Range["B3"].CellStyle.Font.Italic = true;
        formSheet.Range["B3"].CellStyle.Font.Color = ExcelKnownColors.Indigo;

        // Set the required form details
        formSheet.Range[$"B5"].Text = "Ticket Date *";
        formSheet.Range[$"B9"].Text = "Customer Name *";
        formSheet.Range[$"B13"].Text = "Email *";
        formSheet.Range[$"B17"].Text = "Phone";
        formSheet.Range[$"B21"].Text = "Channel *";
        formSheet.Range[$"B25"].Text = "Category *";
        formSheet.Range[$"B29"].Text = "Priority *";
        formSheet.Range[$"B33"].Text = "Requires Follow-up";
        formSheet.Range[$"B37"].Text = "Issue Summary *";
        formSheet.Range[$"B41"].Text = "Details";

        formSheet.SetColumnWidth(2, 28); // B
        formSheet.Range["B1:B41"].CellStyle.Font.Bold = true;
        formSheet.Range["B1:B41"].CellStyle.Font.Size = 14;

        formSheet.SetColumnWidth(4, 40);

        // Set Number Format
        formSheet.Range["D5"].NumberFormat = "dd-mmm-yyyy";
        formSheet.Range["D5"].BorderAround();

        // Add data validation for date
        IDataValidation dataValidation = formSheet.Range["D5"].DataValidation;
        dataValidation.AllowType = ExcelDataType.Date;
        dataValidation.CompareOperator = ExcelDataValidationComparisonOperator.GreaterOrEqual;
        dataValidation.FirstDateTime = new DateTime(2000, 1, 1);
        dataValidation.ErrorBoxText = "Enter a valid date on or after 01-Jan-2000.";

        // Add Form controls
        AddFormControlsAndLinking(formSheet);
    }

    /// <summary>
    /// Add form controls and link to cells
    /// </summary>
    /// <param name="form">Form Worksheet</param>
    private static void AddFormControlsAndLinking(IWorksheet form)
    {
        // Add customer name text box
        ITextBoxShape nameTextBox = form.Shapes.AddTextBox();
        nameTextBox.Left = 330;
        nameTextBox.Top = 200;
        nameTextBox.Height = 30;
        nameTextBox.Width = 280;
        nameTextBox.Name = "CustomerNameTextBox";

        // Add email address text box
        ITextBoxShape emailTxtBox = form.Shapes.AddTextBox();
        emailTxtBox.Left = 330;
        emailTxtBox.Top = 300;
        emailTxtBox.Height = 30;
        emailTxtBox.Width = 280;
        emailTxtBox.Name = "EmailTextBox";

        // Add text box for phone number
        ITextBoxShape phoneTxtBox = form.Shapes.AddTextBox();
        phoneTxtBox.Left = 330;
        phoneTxtBox.Top = 400;
        phoneTxtBox.Height = 30;
        phoneTxtBox.Width = 280;
        phoneTxtBox.Name = "PhoneTextBox";

        // Add combo box for channels
        IComboBoxShape channelComboBox = form.Shapes.AddComboBox();
        channelComboBox.Left = 330;
        channelComboBox.Top = 500;
        channelComboBox.Height = 30;
        channelComboBox.Width = 280;
        channelComboBox.ListFillRange = form.Workbook.Names["Channels"].RefersToRange;
        channelComboBox.LinkedCell = form["G9"];
        channelComboBox.DropDownLines = 6;
        form.Range["H9"].Formula = "=IF(G9>0, INDEX(Channels, G9), \"\")";
        form.Range["G9:H9"].CellStyle.Font.Color = ExcelKnownColors.White;
        form.Workbook.Names.Add("SelectedChannel").RefersToRange = form.Range["H9"];
        channelComboBox.Name = "ChannelComboBox";

        // Add combo box for categories
        IComboBoxShape categoryComboBox = form.Shapes.AddComboBox();
        categoryComboBox.Left = 330;
        categoryComboBox.Top = 600;
        categoryComboBox.Height = 30;
        categoryComboBox.Width = 280;
        categoryComboBox.ListFillRange = form.Workbook.Names["CategoryList"].RefersToRange;
        categoryComboBox.LinkedCell = form["G10"];
        categoryComboBox.DropDownLines = 6;
        form.Range["H10"].Formula = "=IF(AND(G10>0,COUNTA(CategoryList)>0), INDEX(CategoryList, G10), \"\")";
        form.Range["G10:H10"].CellStyle.Font.Color = ExcelKnownColors.White;
        categoryComboBox.Name = "CategoryComboBox";

        // Add priority option buttons
        IOptionButtonShape optLow = form.OptionButtons.AddOptionButton();
        optLow.Left = 330;
        optLow.Top = 700;
        optLow.Height = 30;
        optLow.Width = 150;
        optLow.Text = "Low";
        optLow.Name = "Priority";

        IOptionButtonShape optMedium = form.OptionButtons.AddOptionButton();
        optMedium.Left = 480;
        optMedium.Top = 700;
        optMedium.Height = 30;
        optMedium.Width = 150;
        optMedium.Text = "Medium";
        optMedium.Name = "Priority";

        IOptionButtonShape optHigh = form.OptionButtons.AddOptionButton();
        optHigh.Left = 630;
        optHigh.Top = 700;
        optHigh.Height = 30;
        optHigh.Width = 150;
        optHigh.Text = "High";
        optHigh.Name = "Priority";

        IOptionButtonShape optCritical = form.OptionButtons.AddOptionButton();
        optCritical.Left = 780;
        optCritical.Top = 700;
        optCritical.Height = 30;
        optCritical.Width = 150;
        optCritical.Text = "Critical";
        optCritical.LinkedCell = form["G11"];
        optCritical.Name = "Priority";
        optMedium.CheckState = ExcelCheckState.Checked;
        form.Range["H11"].Formula = "=IF(G11>0, INDEX(Priorities, G11), \"\")";
        form.Range["G11:H11"].CellStyle.Font.Color = ExcelKnownColors.White;


        // Add checkbox for follow up
        ICheckBoxShape followCb = form.Shapes.AddCheckBox();
        followCb.Left = 330;
        followCb.Top = 800;
        followCb.Height = 30;
        followCb.Width = 200;
        followCb.LinkedCell = form["G12"];
        followCb.Text = "Requires Follow-up";
        followCb.CheckState = ExcelCheckState.Unchecked;
        followCb.Line.ForeColor = Color.White;
        form.Range["H12"].Formula = "=IF(G12, \"Yes\", \"No\")";
        form.Range["G12:H12"].CellStyle.Font.Color = ExcelKnownColors.White;
        followCb.Name = "FollowUpCheckBox";

        // Add summary text box
        ITextBoxShape summaryTb = form.Shapes.AddTextBox();
        summaryTb.Left = 330;
        summaryTb.Top = 900;
        summaryTb.Height = 30;
        summaryTb.Width = 320;
        summaryTb.Name = "IssueSummaryTextBox";

        // Add details text box
        ITextBoxShape detailsTb = form.Shapes.AddTextBox();
        detailsTb.Left = 330;
        detailsTb.Top = 1000;
        detailsTb.Height = 100;
        detailsTb.Width = 320;
        detailsTb.Name = "DetailsTextBox";

        // Hide the gridlines
        form.IsGridLinesVisible = false;
    }
}
