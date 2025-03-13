import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

function App() {
  const generateExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Demo Sheet");
    const worksheet2 = workbook.addWorksheet("Demo Sheet 2");

    // Define columns
    worksheet.columns = [
      { header: "Textual col", key: "name", width: 20 },
      { header: "Numeric col", key: "score", width: 10 },
      { header: "Percentage col", key: "percentage", width: 15 },
      { header: "Currency col", key: "salary", width: 15 },
      { header: "Date col", key: "date", width: 15 },
    ];

    // Add sample data
    worksheet.addRow({
      name: "Test 1",
      score: 95,
      percentage: 0.95,
      salary: 50000,
      date: new Date(),
    });
    worksheet.addRow({
      name: "Test 2",
      score: 88,
      percentage: 0.88,
      salary: 60000,
      date: new Date(),
    });

    worksheet2.columns = [
      { header: "Textual col", key: "name", width: 20 },
      { header: "Numeric col", key: "score", width: 10 },
      { header: "Percentage col", key: "percentage", width: 15 },
      { header: "Currency col", key: "salary", width: 15 },
      { header: "Date col", key: "date", width: 15 },
    ];

    // Add sample data
    worksheet2.addRow({
      name: "Test 3",
      score: 95,
      percentage: 0.95,
      salary: 50000,
      date: new Date(),
    });
    worksheet2.addRow({
      name: "Test 4",
      score: 88,
      percentage: 0.88,
      salary: 60000,
      date: new Date(),
    });

    // Apply formatting
    worksheet.getColumn("percentage").numFmt = "0.00%";
    worksheet.getColumn("salary").numFmt = '"$"#,##0.00';
    worksheet2.getColumn("percentage").numFmt = "0.00%";
    worksheet2.getColumn("salary").numFmt = '"$"#,##0.00';

    // Generate Excel file
    const buffer = await workbook.xlsx.writeBuffer();

    // Trigger download
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    saveAs(blob, "exported_data.xlsx");
  };

  return (
    <div style={{ textAlign: "center", marginTop: "50px" }}>
      <h1>React + ExcelJS Demo</h1>
      <button onClick={generateExcel}>Download Excel</button>
    </div>
  );
}

export default App;
