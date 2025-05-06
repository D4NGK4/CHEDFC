function registerTemplate(templateName, personnelInitials) {
  const spreadsheetId = "1hm5BMdt7L9P_A9vZ5WGFYL_77KYnXBUTtZ8HihCfofk";
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const formSheet = spreadsheet.getSheetByName("Form Responses 1");
    const counterSheet = spreadsheet.getSheetByName("Counter");

    // Get current year
    const currentYear = new Date().getFullYear();
    const yearShort = currentYear.toString().slice(-2);

    // Get all existing form responses
    const formDataValues = formSheet.getDataRange().getValues();

    // Get headers to find if personalities column exists
    const headers = formDataValues[0];
    let personalitiesColumnIndex = headers.indexOf("Personalities");

    // If the column doesn't exist, add it
    if (personalitiesColumnIndex === -1) {
      personalitiesColumnIndex = headers.length;
      formSheet
        .getRange(1, personalitiesColumnIndex + 1)
        .setValue("Personalities");
    }

    // Find existing row with same template ID and year
    const existingRowIndex = formDataValues.findIndex(
      (row, index) =>
        index > 0 && row[0] === templateName && row[1] === currentYear
    );

    let count;

    if (existingRowIndex !== -1) {
      // Increment only the count
      const existingRow = formDataValues[existingRowIndex];
      count = existingRow[2] + 1;

      // Update count and personalities
      formSheet.getRange(existingRowIndex + 1, 3).setValue(count);
      formSheet
        .getRange(existingRowIndex + 1, personalitiesColumnIndex + 1)
        .setValue(personnelInitials);
    } else {
      // No existing entry - create new with count 1
      count = 1;
      const controlNumber = yearShort + "-" + ("0000" + count).slice(-4);
      const templateData = mapTemplateData(templateName);

      // Create an array with empty values for all columns up to the personalities column
      const rowData = new Array(Math.max(personalitiesColumnIndex + 1, 5));

      // Fill in the values we know
      rowData[0] = templateName;
      rowData[1] = currentYear;
      rowData[2] = count;
      rowData[3] = controlNumber;

      if (templateData && templateData["Fields"]) {
        rowData[4] = templateData["Fields"];
      }

      rowData[personalitiesColumnIndex] = personnelInitials;

      formSheet.appendRow(rowData);
    }

    // Update counter sheet
    const counterData = counterSheet.getDataRange().getValues();
    let counterRow = counterData.findIndex(
      (row) => row[0] === templateName && row[1] === currentYear
    );

    if (counterRow === -1) {
      counterSheet.appendRow([templateName, currentYear, count]);
    } else {
      counterSheet.getRange(counterRow + 1, 3).setValue(count);
    }

    return {
      status: "success",
      count: count,
      templateName: templateName,
      fileUrl: `https://docs.google.com/document/d/${templateName}/edit`,
    };
  } catch (error) {
    Logger.log("Error in registerTemplate: " + error.toString());
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function submitFormData(templateName) {
  const spreadsheetId = "1hm5BMdt7L9P_A9vZ5WGFYL_77KYnXBUTtZ8HihCfofk";
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const formSheet = spreadsheet.getSheetByName("Form Responses 1");
    const counterSheet = spreadsheet.getSheetByName("Counter");

    // Get current year
    const currentYear = new Date().getFullYear();
    const yearShort = currentYear.toString().slice(-2);

    // Get all existing form responses
    const formDataValues = formSheet.getDataRange().getValues();

    // Find existing row with same template ID and year
    const existingRowIndex = formDataValues.findIndex(
      (row) => row[0] === templateName && row[1] === currentYear
    );

    let count, finalControlNumber;

    if (existingRowIndex !== -1) {
      const existingRow = formDataValues[existingRowIndex];
      count = existingRow[2] + 1;
      finalControlNumber = yearShort + "-" + ("0000" + count).slice(-4);

      formSheet
        .getRange(existingRowIndex + 1, 3, 1, 2)
        .setValues([[count, finalControlNumber]]);
    } else {
      // No existing entry - create new with count 1
      count = 1;
      finalControlNumber = yearShort + "-" + ("0000" + count).slice(-4);

      const templateData = mapTemplateData(templateName);

      formSheet.appendRow([
        templateName,
        currentYear,
        count,
        finalControlNumber,
        templateData["Fields"],
      ]);
    }

    // Update counter sheet
    const counterData = counterSheet.getDataRange().getValues();
    let counterRow = counterData.findIndex(
      (row) => row[0] === templateName && row[1] === currentYear
    );

    if (counterRow === -1) {
      counterSheet.appendRow([templateName, currentYear, count]);
    } else {
      counterSheet.getRange(counterRow + 1, 3).setValue(count);
    }

    return {
      status: "success",
      controlNumber: finalControlNumber,
      fileUrl: `https://docs.google.com/document/d/${templateName}/edit`,
    };
  } catch (error) {
    Logger.log("Error in submitFormData: " + error.toString());
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function cancelFormRequest(controlNumber) {
  const spreadsheetId = "1hm5BMdt7L9P_A9vZ5WGFYL_77KYnXBUTtZ8HihCfofk";

  if (controlNumber.startsWith("TEMP-")) {
    return { status: "removed", controlNumber: controlNumber };
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const formSheet = spreadsheet.getSheetByName("Form Responses 1");
    const data = formSheet.getDataRange().getValues();

    // Find the row with matching control number
    const rowIndex = data.findIndex((row) => row[3] === controlNumber);

    if (rowIndex !== -1) {
      const templateId = data[rowIndex][0];
      const year = data[rowIndex][1];
      const currentCount = data[rowIndex][2];

      // Only delete if count is 1, otherwise decrement
      if (currentCount <= 1) {
        formSheet.deleteRow(rowIndex + 1);
      } else {
        // Decrement count and update control number
        const newCount = currentCount - 1;
        const yearShort = year.toString().slice(-2);
        const newControlNumber =
          yearShort + "-" + ("0000" + newCount).slice(-4);

        formSheet
          .getRange(rowIndex + 1, 3, 1, 2)
          .setValues([[newCount, newControlNumber]]);
      }

      // Update counter sheet
      const counterSheet = spreadsheet.getSheetByName("Counter");
      const counterData = counterSheet.getDataRange().getValues();
      const counterRow = counterData.findIndex(
        (row) => row[0] === templateId && row[1] === year
      );

      if (counterRow !== -1) {
        counterSheet.getRange(counterRow + 1, 3).setValue(currentCount - 1);
      }

      return { status: "removed", controlNumber: controlNumber };
    }

    return { status: "not_found", controlNumber: controlNumber };
  } catch (error) {
    Logger.log("Error in cancelFormRequest: " + error.toString());
    throw error;
  }
}
