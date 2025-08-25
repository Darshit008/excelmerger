const XLSX = require("xlsx");
const fs = require("fs");

function mergeSheets(inputFile, outputFile = "merged_sheets.xlsx") {
  try {
    console.log("📁 Reading file:", inputFile);

    // Load workbook
    const workbook = XLSX.readFile(inputFile);

    // Get all sheet names
    const sheetNames = workbook.SheetNames;
    console.log("📑 Found sheets:", sheetNames);

    let allData = [];
    let allColumns = new Set();

    // Read all sheets
    sheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" }); // defval="" => empty cells as ""
      console.log(`   Reading sheet: ${sheetName}, Rows: ${jsonData.length}`);

      if (jsonData.length > 0) {
        // Track all unique columns
        Object.keys(jsonData[0]).forEach((col) => allColumns.add(col));
        // Add Source_Sheet column
        jsonData.forEach((row) => (row["Source_Sheet"] = sheetName));
        allData = allData.concat(jsonData);
      }
    });

    if (allData.length === 0) {
      console.log("❌ No data found in any sheet!");
      return;
    }

    // Final columns order
    const finalColumns = Array.from(allColumns);
    finalColumns.push("Source_Sheet");

    // Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(allData, { header: finalColumns });

    // Create new workbook and append sheet
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, worksheet, "Merged");

    // Save file
    XLSX.writeFile(newWorkbook, outputFile);
    console.log("✅ Merged sheets saved to:", outputFile);
  } catch (err) {
    console.error("❌ Error:", err.message);
  }
}

// Example usage
mergeSheets("File.xlsx", "merged_sheets.xlsx");
