# Create an Excel Workbook in a React App Using the JavaScript Excel Library 

This sample demonstrates how to create an Excel workbook using the `Syncfusion JavaScript Excel Library` in a `React` application.

The generated workbook includes a professionally formatted expense report with merged cells, custom fonts, borders, alternating row colors, formulas, date and currency formatting, and automatic file download functionality.

## Features Demonstrated

This sample shows how to:

- Create an Excel workbook programmatically.
- Configure worksheet layouts, row heights, and column widths.
- Merge cells and create report headers.
- Apply fonts, colors, borders, and alignments.
- Format date and currency values.
- Populate worksheet cells with data.
- Add formulas for calculations.
- Create alternating row styles for improved readability.
- Save and download the generated Excel workbook.

## Project Structure

### App.jsx

The `generateExcel()` method performs the following operations:

1. Creates a new Excel workbook.
2. Accesses and configures the active worksheet.
3. Creates a formatted report title.
4. Adds employee information.
5. Applies custom styling to worksheet sections.
6. Populates expense categories and daily expense data.
7. Adds formulas to calculate reimbursements and totals.
8. Applies date and currency formatting.
9. Saves and downloads the workbook as `Output.xlsx`.

### User Interface

The user interface includes:

- A sample description.
- A **Generate Excel** button.
- Instructions for generating the Excel workbook.

When the button is clicked, the workbook is generated and downloaded automatically.

## Prerequisites

Ensure the following software is installed:

- Node.js
- npm

## Installation

Install the Syncfusion JavaScript Excel Library package:

```bash
npm install @syncfusion/ej2-xlsx
```

## Running the Application

Start the Vite development server:

```bash
npm run dev
```

Then open:

```text
http://localhost:5173
```

Click **Generate Excel** to create and download the Excel workbook.

## Additional Resources

### Documentation

[JavaScript Excel Library Documentation](https://help.syncfusion.com/document-processing/excel/excel-library/javascript/create-excel-document-react )

### Online Demos

[JavaScript Excel Library Demos](https://document.syncfusion.com/demos/document-library/react/#/tailwind3/excel/getting-started )

### Product Page

[JavaScript Excel Library Product Page](https://www.syncfusion.com/document-sdk/javascript-excel-library )