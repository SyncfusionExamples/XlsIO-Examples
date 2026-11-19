import { Workbook, BorderStyle } from '@syncfusion/ej2-xlsx';

function App() {

  // Workbook theme colors.
  const ACCENT = '0070C0';
  const LIGHT_BAND = 'DEEAF6';
  const GRAY_TEXT = '595959';

  // Returns a reusable thin border style.
  const thinEdge = () => ({
    style: "Thin",
    color: { rgb: "B4C6E7" }
  });

  // Applies borders to all four sides of a cell.
  const boxBorder = () => ({
    left: thinEdge(),
    right: thinEdge(),
    top: thinEdge(),
    bottom: thinEdge()
  });

  // Generates and downloads the Excel workbook.
  const generateExcel = async () => {

    // Create a new workbook and get the active worksheet.
    const workbook = Workbook.create();
    const sheet = workbook.activeSheet;

    // Configure worksheet layout.
    sheet.isGridLinesVisible = false;
    sheet.setColumnWidth(1, 30);
    sheet.setColumnWidth(2, 30);
    sheet.setColumnWidth(3, 30);
    sheet.setColumnWidth(4, 30);
    sheet.setRowHeight(2, 40);
    sheet.setRowHeight(9, 24);

    // Create the report title.
    sheet.range('A2:D2').merge();
    sheet.cell('A2').value = 'EXPENSE REPORT';
    sheet.cell('A2').style = {
      font: { name: 'Verdana', bold: true, size: 28, color: { rgb: ACCENT } },
      alignment: { horizontal: 'center', vertical: 'center' }
    };

    // Add employee information.
    sheet.cell('A4').value = 'Employee';
    sheet.cell('B4').value = 'Roger Federer';
    sheet.cell('A5').value = 'Department';
    sheet.cell('B5').value = 'Administration';

    // Format the employee information section.
    for (let row = 4; row <= 7; row++) {
      for (let col = 1; col <= 2; col++) {
        sheet.cell(row, col).style = {
          font: {
            name: 'Verdana', bold: col === 1, size: 11,
            color: { rgb: col === 1 ? GRAY_TEXT : '404040' }
          },
          fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: LIGHT_BAND } },
          alignment: { horizontal: col === 1 ? 'left' : 'right', vertical: 'center' }
        };
      }
    }

    // Add the week ending date.
    sheet.cell('A6').value = 'Week Ending';
    sheet.cell('B6').dateTime = new Date(2012, 9, 20);
    sheet.cell('B6').numberFormat = 'm/d/yyyy';

    // Add the mileage reimbursement rate.
    sheet.cell('A7').value = 'Mileage Rate';
    sheet.cell('B7').number = 0.70;
    sheet.cell('B7').numberFormat = '$#,##0.00';

    // Format the expense table area.
    for (let row = 9; row <= 20; row++) {
      let banded = row > 9 && (row - 9) % 2 === 0;
      for (let col = 1; col <= 4; col++) {
        sheet.cell(row, col).style = {
          font: { name: 'Verdana', size: 11, color: { rgb: GRAY_TEXT } },
          fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: banded ? 'F2F7FC' : 'FFFFFF' } },
          border: boxBorder(),
          alignment: { horizontal: col === 1 ? 'left' : 'right', vertical: 'center' }
        }
      }
    }

    // Add table headers.
    sheet.cell('A9').value = 'Expenses';
    sheet.cell('B9').value = 'Day 1';
    sheet.cell('C9').value = 'Day 2';
    sheet.cell('D9').value = 'Day 3';

    // Apply header styling.
    for (let col = 1; col <= 4; col++) {
      sheet.cell(9, col).style = {
        font: { name: 'Verdana', bold: true, size: 11, color: { rgb: 'FFFFFF' } },
        fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: ACCENT } },
        border: boxBorder(),
        alignment: { horizontal: col === 1 ? 'left' : 'right', vertical: 'center' }
      };
    }

    // Define the expense categories.
    let labels = ['Miles Driven', 'Miles Reimbursement', 'Parking and Tolls', 'Auto Rental',
      'Lodging', 'Breakfast', 'Lunch', 'Dinner', 'Snacks', 'Others', 'Total'];

    // Populate the expense categories.
    for (let i = 0; i < labels.length; i++) {
      sheet.cell(10 + i, 1).value = labels[i];
    }

    // Add Day 1 expense details.
    sheet.cell('B10').number = 100;
    sheet.cell('B11').formula = '=(B7*B10)';
    let day1 = [0, 0, 0, 9, 12, 13, 9.5, 0];
    for (let i = 0; i < day1.length; i++) {
      sheet.cell(12 + i, 2).number = day1[i];
    }

    // Calculate the Day 1 total.
    sheet.cell('B20').formula = '=SUM(B11:B19)';

    // Format Day 1 values as currency.
    for (let row = 11; row <= 20; row++) {
      sheet.cell(row, 2).numberFormat = '$#,##0.00';
    }

    // Add Day 2 expense details.
    sheet.cell('C10').number = 145;
    sheet.cell('C11').formula = '=(B7*C10)';
    let day2 = [15, 0, 45, 9, 12, 15, 7, 0];
    for (let i = 0; i < day2.length; i++) {
      sheet.cell(12 + i, 3).number = day2[i];
    }

    // Calculate the Day 2 total.
    sheet.cell('C20').formula = '=SUM(C11:C19)';

    // Format Day 2 values as currency.
    for (let row = 11; row <= 20; row++) {
      sheet.cell(row, 3).numberFormat = '$#,##0.00';
    }

    // Add Day 3 expense details.
    sheet.cell('D10').number = 113;
    sheet.cell('D11').formula = '=(B7*D10)';
    let day3 = [17, 8, 45, 7, 11, 16, 7, 5];
    for (let i = 0; i < day3.length; i++) {
      sheet.cell(12 + i, 4).number = day3[i];
    }

    // Calculate the Day 3 total.
    sheet.cell('D20').formula = '=SUM(D11:D19)';
    for (let row = 11; row <= 20; row++) {
      sheet.cell(row, 4).numberFormat = '$#,##0.00';
    }

    // Create the total row.
    sheet.cell(20, 1).value = 'Total';

    // Highlight the total row.
    for (let col = 1; col <= 4; col++) {
      sheet.cell(20, col).style = {
        font: { name: 'Verdana', bold: true, size: 11, color: { rgb: 'FFFFFF' } },
        fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: ACCENT } },
        border: {
          left: { style: BorderStyle.Thin, color: { rgb: ACCENT } },
          right: { style: BorderStyle.Thin, color: { rgb: ACCENT } },
          top: { style: BorderStyle.Medium, color: { rgb: ACCENT } },
          bottom: { style: BorderStyle.Medium, color: { rgb: ACCENT } }
        },
        alignment: { horizontal: col === 1 ? 'left' : 'right', vertical: 'center' }
      };
    }

    // Save the workbook.
    const bytes = await workbook.save();

    // Create a blob from the workbook bytes.
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    // Create a download URL.
    const url = URL.createObjectURL(blob);

    // Create a temporary download link.
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'Output.xlsx';

    // Download the generated workbook.
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();

    // Release the object URL.
    URL.revokeObjectURL(url);
  }

  // Render the application UI.
  return (
    <div
      style={{
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        minHeight: "100vh",
        backgroundColor: "#f4f7fc",
        padding: "20px",
      }}
    >
      <div
        style={{
          width: "600px",
          backgroundColor: "#ffffff",
          padding: "32px",
          borderRadius: "12px",
          boxShadow: "0 4px 16px rgba(0,0,0,0.1)",
          textAlign: "center",
        }}
      >
        <h2
          style={{
            color: "#0f6cbd",
            margin: "0 0 16px 0",
            fontSize: "32px",
            fontWeight: "700",
          }}
        >
          JavaScript Excel Library
        </h2>

        <p
          style={{
            margin: "0",
            fontSize: "16px",
            lineHeight: "1.7",
            color: "#555",
          }}
        >
          Click the button below to generate an Excel workbook using the
          Syncfusion JavaScript Excel Library.
        </p>

        <button
          onClick={generateExcel}
          style={{
            marginTop: "24px",
            backgroundColor: "#0f6cbd",
            color: "#ffffff",
            border: "none",
            padding: "12px 30px",
            fontSize: "16px",
            fontWeight: "600",
            borderRadius: "6px",
            cursor: "pointer",
          }}
        >
          Generate Excel
        </button>
      </div>
    </div>
  )
}

export default App
