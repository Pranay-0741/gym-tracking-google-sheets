function handleMonthlyGymAutomation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateName = 'BaseTemplate';

  const now = new Date();
  const monthSheetName = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'MMMM yyyy');

  // === Step 1: Create Monthly Sheet if Needed ===
  let monthSheet = ss.getSheetByName(monthSheetName);
  if (!monthSheet) {
    const templateSheet = ss.getSheetByName(templateName);
    if (!templateSheet) throw new Error(`Template sheet "${templateName}" not found.`);
    monthSheet = templateSheet.copyTo(ss).setName(monthSheetName);
    ss.setActiveSheet(monthSheet);
    ss.moveActiveSheet(ss.getNumSheets());
    Logger.log(`Created new monthly sheet: ${monthSheetName}`);
  }

  // === Step 2: Transfer Monthly Data from _RawResponses ===
  const rawSheet = ss.getSheetByName('_RawResponses');
  const rawData = rawSheet.getDataRange().getValues();
  const headers = rawData[0];
  const data = rawData.slice(1);

  const timestampCol = headers.indexOf("Timestamp");
  if (timestampCol === -1) throw new Error("Timestamp column not found.");

  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  const targetTimestamps = monthSheet.getRange("A7:A1000").getValues().flat();
  const alreadyWritten = targetTimestamps.filter(v => v instanceof Date).map(v => v.toString());

  const rowsToWrite = [];

  data.forEach(row => {
    const ts = row[timestampCol];
    if (ts instanceof Date && ts.getMonth() === currentMonth && ts.getFullYear() === currentYear) {
      if (!alreadyWritten.includes(ts.toString())) {
        rowsToWrite.push(row.slice(0, 9)); // A to I
      }
    }
  });

  if (rowsToWrite.length > 0) {
    let insertRow = 7;
    const existingData = monthSheet.getRange("A7:A1000").getValues().flat();
    for (let i = 0; i < existingData.length; i++) {
      if (!existingData[i]) {
        insertRow = 7 + i;
        break;
      }
    }
    monthSheet.getRange(insertRow, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
    Logger.log(`${rowsToWrite.length} new rows added to ${monthSheetName}`);
  } else {
    Logger.log("No new rows to transfer.");
  }

  // === Step 3: Duplicate Check ===
  const checkRange = monthSheet.getRange("A7:N1000");
  const checkData = checkRange.getValues();

  const headerRow = monthSheet.getRange("A6:N6").getValues()[0];
  const nameIndex = headerRow.indexOf("Full Name");
  const phoneIndex = headerRow.indexOf("Phone Number");
  const resultIndex = 13; // Column N

  const seen = [];
  const duplicateResults = [];

  checkData.forEach((row, i) => {
    const name = row[nameIndex];
    const phone = row[phoneIndex];

    if (!name && !phone) {
      duplicateResults.push([""]);
      return;
    }

    let isDuplicate = false;
    let isPartial = false;

    for (const entry of seen) {
      if (entry.name === name && entry.phone === phone) {
        isDuplicate = true;
        break;
      } else if (entry.name === name || entry.phone === phone) {
        isPartial = true;
      }
    }

    if (isDuplicate) {
      duplicateResults.push(["Duplicate"]);
    } else if (isPartial) {
      duplicateResults.push(["Partial Duplicate"]);
    } else {
      duplicateResults.push([""]);
    }

    seen.push({ name: name || "", phone: phone || "" });
  });

  monthSheet.getRange(7, resultIndex + 1, duplicateResults.length, 1).setValues(duplicateResults);
  Logger.log(`Duplicate check completed for ${monthSheetName}`);
}
