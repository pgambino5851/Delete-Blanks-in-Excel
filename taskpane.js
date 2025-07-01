//      <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("deleteRowsButton").onclick = deleteRowsWithAnyBlankCell;
  }
  console.log("Office is ready");
});

function deleteRowsWithAnyBlankCell() {
  Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange();
    selectedRange.load(["address", "rowIndex", "rowCount", "columnIndex", "columnCount", "worksheet"]);
    
    const worksheet = selectedRange.worksheet;
    const usedRange = worksheet.getUsedRange();
    usedRange.load(["address", "rowIndex", "rowCount", "columnIndex", "columnCount"]);
    
    await context.sync();

    console.log("Selected range:", selectedRange.address);
    console.log("Worksheet name:", worksheet.name);
    console.log("Used range:", usedRange.address);

    // Determine if full sheet is selected (Ctrl+A)
    const isFullSheetSelected =
      selectedRange.rowCount >= usedRange.rowCount &&
      selectedRange.columnCount >= usedRange.columnCount;

    let rangeToCheck;

    if (isFullSheetSelected) {
      console.log("Full sheet selected. Using usedRange instead.");
      rangeToCheck = usedRange;
    } else {
      const checkColumnCount = Math.min(selectedRange.columnCount, usedRange.columnCount);
      rangeToCheck = worksheet.getRangeByIndexes(
        selectedRange.rowIndex,
        selectedRange.columnIndex,
        selectedRange.rowCount,
        checkColumnCount
      );
    }

    rangeToCheck.load(["address", "values", "rowIndex", "rowCount"]);
    await context.sync();

    console.log("Range to check address:", rangeToCheck.address);

    const values = rangeToCheck.values;
    if (!values || values.length === 0) {
      console.log("No data found in selection.");
      return;
    }

    const rowsToDelete = [];

    // Identify rows with any blank cells
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (row.some(cell => cell == null || String(cell).trim() === "")) {
        rowsToDelete.push(i);
      }
    }

    console.log("Rows to delete (within selection):", rowsToDelete);

    // Delete rows from bottom up to avoid index shifting
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      const rowNumber = rangeToCheck.rowIndex + rowsToDelete[i];
      const targetRowRange = worksheet.getRange(`${rowNumber + 1}:${rowNumber + 1}`);
      targetRowRange.delete(Excel.DeleteShiftDirection.up);
    }

    await context.sync();
    console.log(`Deleted ${rowsToDelete.length} rows with blank cells within selection.`);

    if (rowsToDelete.length === 0) {
      console.log("No rows with blank cells found in the selection.");
      document.getElementById("status").textContent = "No rows with blank cells found in the selection.";
      return;
    } else if (rowsToDelete.length === 1) {
      console.log("Deleted 1 row with blank cells within selection.");
      document.getElementById("status").textContent = "Deleted 1 row with blank cells within selection.";
    } else {
      document.getElementById("status").textContent = `Deleted ${rowsToDelete.length} rows with blank cells within selection.`;
      console.log(`Deleted ${rowsToDelete.length} rows with blank cells within selection.`);
    }

  })
  .then(() => {
    console.log("Rows with blank cells deleted successfully.");
  })
  .catch((error) => {
    console.error("Error deleting rows with blank cells:", error);
  });
 
}
// This code defines a function to delete rows with any blank cells in the selected range of an Excel worksheet.






