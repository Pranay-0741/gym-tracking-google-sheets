function createMonthlySheetAndProtect() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateName = 'BaseTemplate';

  const now = new Date();
  const monthYear = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'MMMM yyyy');

  let newSheet = ss.getSheetByName(monthYear);
  if (!newSheet) {
    // Create sheet from template
    const templateSheet = ss.getSheetByName(templateName);
    if (!templateSheet) {
      throw new Error(`Template sheet "${templateName}" not found.`);
    }

    newSheet = templateSheet.copyTo(ss);
    newSheet.setName(monthYear);
    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(ss.getNumSheets());

    Logger.log(`Sheet "${monthYear}" created successfully.`);
  } else {
    Logger.log(`Sheet "${monthYear}" already exists.`);
  }

  // Protect sensitive columns
  const columnsToProtect = ['A','J', 'K', 'L', 'AA', 'AB', 'AC', 'AD', 'AE'];
  const existingProtections = newSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  columnsToProtect.forEach(colLetter => {
    const colIndex = columnLetterToIndex(colLetter);
    const rangeToProtect = newSheet.getRange(7, colIndex, 993); // Rows 7 to 999

    const alreadyProtected = existingProtections.some(p => {
      const pRange = p.getRange();
      return pRange.getA1Notation() === rangeToProtect.getA1Notation();
    });

    if (!alreadyProtected) {
      const protection = rangeToProtect.protect();
      protection.setWarningOnly(true);
      Logger.log(`Protected ${colLetter}7:${colLetter}999`);
    } else {
      Logger.log(`Already protected: ${colLetter}7:${colLetter}999`);
    }
  });
}

// Helper: Convert column letter to number (A → 1, B → 2, ..., AA → 27)
function columnLetterToIndex(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column *= 26;
    column += letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return column;
}

