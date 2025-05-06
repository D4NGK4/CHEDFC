function testCreateDocument() {
  const testCases = [
    ["25-0010", "Liceo de Cagayan University", "WORLD"],
  ];

  const results = testCases.map((data, index) => {
    try {
      const docUrl = DocumentGenerator.createDocument("Test Template", data, false);
      return {
        testCase: index + 1,
        format: data.join(","),
        success: !!docUrl,
        result: docUrl || "Failed",
      };
    } catch (e) {
      return {
        testCase: index + 1,
        format: data.join(","),
        success: false,
        error: e.message
      };
    }
  });

  Logger.log("Test Results:\n" + JSON.stringify(results, null, 2));
  return results;
}

function TestGetAllTemplates() {

  const results = DocumentGenerator.getTemplates();

  Logger.log("Test Results:\n" + JSON.stringify(results, null, 2));
  return results;
}

function TestGetATemplate() {

  const results = DocumentGenerator.getTemplate("Training Request Form");

  Logger.log("Test Results:\n" + JSON.stringify(results, null, 2));
  return results;
}

function getFersons() {
  const results = DocumentGenerator.getPersonalities();
  Logger.log("Test Results:\n" + JSON.stringify(results, null, 2));
  Logger.log(results);
  return results;
}

function getStatus() {
  const results = DocumentGenerator.getRequestStatus("1h550v3Odk6uNAoNCnb-yvdbXzAMDZGIRJik_-fD6Wq8");
  Logger.log(results);
  return results;
}

function getInitials() {
  const results = DocumentGenerator.requestStampInitial("clydegevero14@gmail.com", "1Hczs7rW9DrczGMD-vjzWcLr-XA2I3PeZp47khhPdjY4");
  Logger.log("Test Results:\n" + JSON.stringify(results, null, 2));
  return results;
}

function getSignature() {
  const results = DocumentGenerator.requestStampSignature("jmpalasan@ched.gov.ph", "1pZRyDM3TPoN7P7uYhT11VVYIx_CTn4ZA22kzYp51-lo");
  Logger.log(results);
  return results;
}

function setupTrackingSpreadsheet() {
  try {

    const ss = SpreadsheetApp.openById('10kqLkE4WaPZGmpLExi6lO7lul-G8xJ6ONsl5C7k8aCE');
    const sheet = ss.getSheets()[0];

    sheet.getRange("A1:C1").setValues([["FileId", "Email", "Status"]]);

    const headerRange = sheet.getRange("A1:C1");
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f3f3f3");

    sheet.autoResizeColumns(1, 3);

    sheet.setFrozenRows(1);

    return {
      status: "success",
      message: "Tracking spreadsheet has been set up successfully"
    };
  } catch (e) {
    Logger.log("Error setting up tracking spreadsheet: " + e.toString());
    return {
      status: "error",
      message: e.message || "Unknown error occurred"
    };
  }
}

function testTrackSubmission() {
  try {

    const fileId = "1Hczs7rW9DrczGMD-vjzWcLr-XA2I3PeZp47khhPdjY4"; 
    const personalityName = "Clyde Gevero"; 

    DocumentService.trackSubmission(fileId, personalityName);

    DocumentService.updateSubmissionStatus(fileId);

    return {
      status: "success",
      message: "Test submission tracked successfully"
    };
  } catch (e) {
    Logger.log("Error in test track submission: " + e.toString());
    return {
      status: "error",
      message: e.message || "Unknown error occurred"
    };
  }
}

function testRefreshAllStatuses() {
  return DocumentService.refreshAllStatuses();
}

function debugDocumentCreation() {
  try {
    const templateName = "Test Template";
    const data = ["TEST-12345", "Example University", "Example Data"];
    const isPrivate = false;

    Logger.log("Starting document creation debug test...");

    const result = DocumentGenerator.createDocument(templateName, data, isPrivate);

    Logger.log("Document creation result type: " + typeof result);
    Logger.log("Document creation result: " + JSON.stringify(result));

    if (typeof result === 'object') {

      for (const key in result) {
        Logger.log("Property '" + key + "': " + JSON.stringify(result[key]));
      }
    }

    const fileId = DocumentService.extractFileId(result);
    Logger.log("Extracted file ID: " + fileId);

    return {
      status: "success",
      resultType: typeof result,
      result: result,
      extractedFileId: fileId
    };
  } catch (e) {
    Logger.log("Debug document creation error: " + e.toString());
    Logger.log("Stack: " + (e.stack || "No stack available"));

    return {
      status: "error",
      message: e.message || "Unknown error occurred",
      errorDetails: {
        name: e.name || "Error",
        stack: e.stack || "No stack trace available"
      }
    };
  }
}

function debugGetRequestStatus() {
  try {
    const fileId = "1HQzI9rXoPuFdMEH9v6H259B22SlhstXMkVLtIYLajoc";

    Logger.log("---------------------- DEBUG GET REQUEST STATUS ----------------------");
    Logger.log("Testing getRequestStatus for file ID: " + fileId);

    const result = DocumentGenerator.getRequestStatus(fileId);

    Logger.log("Result type: " + typeof result);
    Logger.log("Is array: " + Array.isArray(result));

    try {
      Logger.log("Result as JSON: " + JSON.stringify(result));
    } catch (jsonErr) {
      Logger.log("Error stringifying result: " + jsonErr.toString());
    }

    if (Array.isArray(result)) {
      Logger.log("Array length: " + result.length);

      for (let i = 0; i < result.length; i++) {
        Logger.log("Entry " + i + ":");
        Logger.log("  Type: " + typeof result[i]);
        Logger.log("  Is array: " + Array.isArray(result[i]));

        if (Array.isArray(result[i])) {
          Logger.log("  Entry length: " + result[i].length);

          for (let j = 0; j < result[i].length; j++) {
            Logger.log("  Element [" + i + "][" + j + "]: " + JSON.stringify(result[i][j]));

            if (typeof result[i][j] === 'string' &&
              (result[i][j].includes('<') || result[i][j].includes('>'))) {
              Logger.log("  Element [" + i + "][" + j + "] contains HTML");
            }
          }
        } else {

          Logger.log("  Value: " + JSON.stringify(result[i]));
        }
      }
    } else if (typeof result === 'object') {

      Logger.log("Object properties:");

      for (const key in result) {
        if (result.hasOwnProperty(key)) {
          Logger.log("  Property '" + key + "': " + JSON.stringify(result[key]));
        }
      }
    }

    const status = extractLatestStatus(result);
    Logger.log("Extracted status: " + status);

    if (Array.isArray(result) && result.length > 0) {
      if (Array.isArray(result[0]) && result[0].length > 2) {
        Logger.log("Direct access to result[0][2]: " + result[0][2]);
      }

      if (result.length > 1 && Array.isArray(result[1]) && result[1].length > 2) {
        Logger.log("Direct access to result[1][2]: " + result[1][2]);
      }
    }

    Logger.log("---------------------- END DEBUG ----------------------");

    return {
      status: "success",
      fileId: fileId,
      resultType: typeof result,
      isArray: Array.isArray(result),
      extractedStatus: status,
      rawResult: result
    };
  } catch (e) {
    Logger.log("Error in debugGetRequestStatus: " + e.toString());
    Logger.log("Stack: " + e.stack);

    return {
      status: "error",
      message: e.message || "Unknown error occurred",
      errorDetails: {
        name: e.name || "Error",
        stack: e.stack || "No stack trace available"
      }
    };
  }
}

function trashTemplates() {
  const v = "1";
  const rows = [789,793,794];
  const headers = [
    "Row",
    "Email",
    "File Name",
    "File ID",
    "Document Template Name",
    "Fields",
    "Values",
    "isPrivate",
    "isTrashed",
    "Timestamp"
  ];

  rows.forEach(row => {
    DocumentGenerator.trash(v, row, headers);
  });

  fetchTemplatesTest();
}


function getFersonstoSheet() {
  const personalities = DocumentGenerator.getPersonalities();
  const sheet = SpreadsheetApp.openById('1DoEWGWW9Bfh7W-R7rVxGVyRlV55SHCUKpwEG0mwBA8A')
    .getSheetByName('Personalities');

  sheet.clear();

  if (!personalities || !personalities.values || personalities.values.length === 0) {
    sheet.getRange(1, 1).setValue('No results found.');
    return;
  }

  const headers = personalities.headers || [];
  const dataRows = personalities.values || [];

  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  }
}


function fetchTemplatesTest() {
  const results = DocumentGenerator.getDataDocuments();
  const sheet = SpreadsheetApp.openById('1DoEWGWW9Bfh7W-R7rVxGVyRlV55SHCUKpwEG0mwBA8A').getSheetByName('TempLog');

  sheet.clear();

  if (!results) {
    sheet.getRange(1, 1).setValue('No results found.');
    return;
  }

  if (Array.isArray(results)) {
    if (results.length === 0) {
      sheet.getRange(1, 1).setValue('No templates found.');
      return;
    }

    const headers = results[0].headers || [];
    const dataRows = results[0].displayValues || [];

    if (headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    if (dataRows.length > 0) {
      sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    }
  } else if (typeof results === 'object') {
    const headers = results.headers || [];
    const dataRows = results.displayValues || [];

    if (headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    if (dataRows.length > 0) {
      sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    }
  } else {
    sheet.getRange(1, 1).setValue('Unexpected results format.');
  }
}

function fetchTemplates() {
  try {
    const results = DocumentGenerator.getDataDocuments()
    if (Array.isArray(results) && results.length > 0) {
      return {
        headers: results[0],
        data: results.slice(1)
      };
    } else {
      return {
        headers: [],
        data: []
      };
    }
  } catch (error) {
    console.error("Error fetching templates:", error);
    return {
      error: error.toString(),
      headers: [],
      data: []
    };
  }
}

function getEmailQuota() {
  const emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log(`Remaining email quota: ${emailQuotaRemaining}`);
}