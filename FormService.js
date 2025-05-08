// Helper function to get Division Chief's email based on author's division
function getDivisionChiefEmailByAuthorDivision(authorEmail) {
  try {
    Logger.log("[Division Chief Lookup] Starting for author: " + authorEmail);

    if (!authorEmail) {
      Logger.log("[Division Chief Lookup] No author email provided");
      return null;
    }

    // Get all personalities
    const personalities = DocumentGenerator.getPersonalities();
    if (!personalities || !personalities.values) {
      Logger.log("[Division Chief Lookup] No personalities data found");
      return null;
    }

    // Find author's personality record (email is at index 0)
    const authorRecord = personalities.values.find(
      (person) =>
        person.length > 0 &&
        person[0] &&
        person[0].trim() === authorEmail.trim()
    );

    if (!authorRecord) {
      Logger.log(
        "[Division Chief Lookup] Author record not found for: " + authorEmail
      );
      return null;
    }

    // Division should be at index 5 (after level which is at index 4)
    const authorDivision = authorRecord.length > 5 ? authorRecord[5] : null;
    Logger.log("[Division Chief Lookup] Author's division: " + authorDivision);

    // Find all Division Chiefs
    const divisionChiefs = personalities.values.filter(
      (person) =>
        person.length > 4 && person[4] && person[4].trim() === "Division Chief"
    );

    if (!divisionChiefs || divisionChiefs.length === 0) {
      Logger.log("[Division Chief Lookup] No Division Chiefs found");
      return null;
    }

    // Default to return all chiefs if we can't determine the correct one
    if (!authorDivision) {
      Logger.log(
        "[Division Chief Lookup] No division specified for author, returning all Division Chiefs"
      );
      return divisionChiefs.map((chief) => chief[0]);
    }

    // Match Division Chiefs by division
    let matchingChiefs = [];

    // Case: Author is in Technical Division
    if (authorDivision.includes("Technical")) {
      matchingChiefs = divisionChiefs.filter(
        (chief) =>
          chief.length > 5 &&
          chief[5] &&
          (chief[5].includes("Technical") || chief[5].includes("Both"))
      );
      Logger.log(
        "[Division Chief Lookup] Matching Technical Division Chiefs: " +
          matchingChiefs.length
      );
    }
    // Case: Author is in Administrative Division
    else if (authorDivision.includes("Administrative")) {
      matchingChiefs = divisionChiefs.filter(
        (chief) =>
          chief.length > 5 &&
          chief[5] &&
          (chief[5].includes("Administrative") || chief[5].includes("Both"))
      );
      Logger.log(
        "[Division Chief Lookup] Matching Administrative Division Chiefs: " +
          matchingChiefs.length
      );
    }
    // Case: Author is in Both Divisions
    else if (authorDivision.includes("Both")) {
      // For "Both", include all Division Chiefs
      matchingChiefs = divisionChiefs;
      Logger.log(
        "[Division Chief Lookup] Author is in Both divisions, returning all Division Chiefs"
      );
    }

    if (matchingChiefs.length === 0) {
      Logger.log(
        "[Division Chief Lookup] No matching Division Chiefs found, returning all"
      );
      return divisionChiefs.map((chief) => chief[0]);
    }

    // Extract emails of matching Chiefs
    const chiefEmails = matchingChiefs.map((chief) => chief[0]);
    Logger.log(
      "[Division Chief Lookup] Matched Division Chief emails: " +
        chiefEmails.join(", ")
    );

    return chiefEmails;
  } catch (e) {
    Logger.log("[Division Chief Lookup] ERROR: " + e.toString());
    return null;
  }
}

// Helper function to get emails by level
function getPersonalityEmailsByLevel(level, authorEmail) {
  try {
    const personalities = DocumentGenerator.getPersonalities();
    if (!personalities || !personalities.values) {
      Logger.log("[Get Emails By Level] No personalities data found");
      return [];
    }

    // Special handling for Division Chief
    if (level.trim() === "Division Chief" && authorEmail) {
      const divisionChiefEmails =
        getDivisionChiefEmailByAuthorDivision(authorEmail);
      if (divisionChiefEmails && divisionChiefEmails.length > 0) {
        Logger.log(
          `[Get Emails By Level] Found Division Chief emails based on author's division: ${divisionChiefEmails.join(
            ", "
          )}`
        );
        return divisionChiefEmails;
      }
      // If we couldn't get division-specific chiefs, fall back to the default behavior
      Logger.log(
        "[Get Emails By Level] Falling back to default Division Chief lookup"
      );
    }

    // Find all personalities with the given level
    const matchingPersonalities = personalities.values.filter(
      (person) =>
        person.length > 4 && person[4] && person[4].trim() === level.trim()
    );

    // If no matching personalities found
    if (!matchingPersonalities || matchingPersonalities.length === 0) {
      Logger.log(
        `[Get Emails By Level] No personalities found with level: ${level}`
      );
      return [];
    }

    // Extract emails (index 0)
    const emails = matchingPersonalities.map((person) => person[0]);
    Logger.log(
      `[Get Emails By Level] Found ${
        emails.length
      } emails for level ${level}: ${emails.join(", ")}`
    );
    return emails;
  } catch (e) {
    Logger.log(`[Get Emails By Level] ERROR: ${e.toString()}`);
    return [];
  }
}

function registerTemplate(
  templateName,
  personnelInitials,
  includeAuthorFlag,
  formData,
  includeControlNumberFlag
) {
  const spreadsheetId = "1hm5BMdt7L9P_A9vZ5WGFYL_77KYnXBUTtZ8HihCfofk";
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    Logger.log("[Register Template] Starting with template: " + templateName);
    Logger.log("[Register Template] Level Order: " + personnelInitials);
    Logger.log("[Register Template] Include Author Flag: " + includeAuthorFlag);
    Logger.log(
      "[Register Template] Include Control Number: " + includeControlNumberFlag
    );
    Logger.log(
      "[Register Template] Form Data provided: " + (formData ? "Yes" : "No")
    );

    // Track whether we want to include author in approval workflow
    const includeAuthor = !!includeAuthorFlag;
    // Track whether we want to include control number
    const includeControlNumber = includeControlNumberFlag !== false;

    Logger.log(
      "[Register Template] Include Author: " + (includeAuthor ? "Yes" : "No")
    );
    Logger.log(
      "[Register Template] Include Control Number: " +
        (includeControlNumber ? "Yes" : "No")
    );

    // Get current user's email
    let currentUserEmail = "";
    try {
      currentUserEmail = Session.getActiveUser().getEmail();
      Logger.log("[Register Template] Current user email: " + currentUserEmail);
    } catch (e) {
      Logger.log(
        "[Register Template] Error getting user email: " + e.toString()
      );
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const formSheet = spreadsheet.getSheetByName("Form Responses 1");
    const counterSheet = spreadsheet.getSheetByName("Counter");

    // Get current year
    const currentYear = new Date().getFullYear();
    const yearShort = currentYear.toString().slice(-2);

    // Get all existing form responses
    const formDataValues = formSheet.getDataRange().getValues();

    // Get headers to find column indices
    const headers = formDataValues[0];
    let queueColumnIndex = headers.indexOf("Queue");
    let fieldsColumnIndex = headers.indexOf("Fields");
    let fileIdColumnIndex = headers.indexOf("File ID");

    // If the Queue column doesn't exist, add it
    if (queueColumnIndex === -1) {
      // The Queue column should be at column F (index 5)
      queueColumnIndex = 5;
      formSheet.getRange(1, queueColumnIndex + 1).setValue("Queue");
    }

    // If the File ID column doesn't exist, add it after Queue
    if (fileIdColumnIndex === -1) {
      // File ID column should be at column G (index 6, which is queueColumnIndex + 1)
      fileIdColumnIndex = queueColumnIndex + 1;
      formSheet.getRange(1, fileIdColumnIndex + 1).setValue("File ID");
    }

    // Get template data to update fields
    const templateData = mapTemplateData(templateName);
    let fields = templateData ? templateData["Fields"] || "" : "";

    // Collect all emails for this registration in a flat array
    let allEmails = [];

    // Start with the author if includeAuthor is true
    if (includeAuthor && currentUserEmail) {
      allEmails.push(currentUserEmail);
    }

    // Process personnel levels if provided
    if (personnelInitials && personnelInitials.trim()) {
      const levels = personnelInitials.split(",").map((level) => level.trim());

      // For each level, get all associated emails and add them to the flat array
      levels.forEach((level) => {
        // Pass the author's email to getPersonalityEmailsByLevel for Division Chief handling
        const levelEmails = getPersonalityEmailsByLevel(
          level,
          currentUserEmail
        );
        if (levelEmails.length > 0) {
          // Add all emails from this level to our flat array
          allEmails = allEmails.concat(levelEmails);
        } else {
          Logger.log(`[Register Template] No emails found for level: ${level}`);
        }
      });
    }

    // Join all emails with commas - this is the new batch of emails for this registration
    const emailBatch = allEmails.join(",");
    Logger.log(
      `[Register Template] Email batch for this registration: ${emailBatch}`
    );

    let count = 0;
    let controlNumber = "";
    let documentFileId = ""; // To store the document file ID
    let documentResult = null; // Initialize documentResult variable

    // Always generate a control number for tracking, regardless of includeControlNumber flag
    // Find existing row with same template ID and year
    const existingRowIndex = formDataValues.findIndex(
      (row, index) =>
        index > 0 && row[0] === templateName && row[1] === currentYear
    );

    Logger.log(
      "[Register Template] Searching for existing template entry: " +
        templateName +
        " for year " +
        currentYear +
        ", found at index: " +
        (existingRowIndex !== -1 ? existingRowIndex : "not found")
    );

    if (existingRowIndex !== -1) {
      // Increment only the count
      const existingRow = formDataValues[existingRowIndex];
      count = existingRow[2] + 1;

      // Generate the new control number based on the incremented count
      controlNumber = yearShort + "-" + ("0000" + count).slice(-4);
      Logger.log(
        "[Register Template] Incrementing count to " +
          count +
          ", new control number: " +
          controlNumber
      );

      // Update count and control number
      formSheet
        .getRange(existingRowIndex + 1, 3, 1, 2)
        .setValues([[count, controlNumber]]);

      // Get existing queue value and append the new batch with semicolon
      let existingQueue = existingRow[queueColumnIndex] || "";
      let updatedQueue = existingQueue;

      // If we have emails for this registration, add them to the queue
      if (emailBatch) {
        if (existingQueue && existingQueue.trim()) {
          // Add semicolon separator between batches
          updatedQueue = existingQueue + "; " + emailBatch;
        } else {
          // First batch, no separator needed
          updatedQueue = emailBatch;
        }
      }

      Logger.log(
        `[Register Template] Updating queue from "${existingQueue}" to "${updatedQueue}"`
      );

      // Update the queue
      formSheet
        .getRange(existingRowIndex + 1, queueColumnIndex + 1)
        .setValue(updatedQueue);

      // Update the fields if available
      if (fields && fieldsColumnIndex !== -1) {
        formSheet
          .getRange(existingRowIndex + 1, fieldsColumnIndex + 1)
          .setValue(fields);
      }

      // Create document if form data was provided
      if (formData && Array.isArray(formData) && formData.length > 0) {
        Logger.log(
          "[Register Template] Creating document with provided form data"
        );
        try {
          // Only pass control number to document creation if includeControlNumber is true
          const docControlNumber = includeControlNumber ? controlNumber : "";
          documentResult = createDoc(formData, docControlNumber, templateName);
          Logger.log(
            "[Register Template] Document created successfully: " +
              JSON.stringify(documentResult)
          );

          // Extract and store file ID
          if (documentResult && documentResult.status === "success") {
            // Extract file ID from documentUrl
            const docUrl = documentResult.documentUrl || "";
            if (docUrl) {
              documentFileId = docUrl.match(/\/d\/([^\/]+)/)?.[1] || "";
              Logger.log(
                `[Register Template] Extracted file ID: ${documentFileId}`
              );

              // Update File ID in the spreadsheet
              if (documentFileId) {
                // Get the current value of the File ID cell
                const currentFileId = formSheet
                  .getRange(existingRowIndex + 1, fileIdColumnIndex + 1)
                  .getValue();
                let updatedFileId = "";

                // If there's already a file ID, append the new one with a comma
                if (currentFileId && currentFileId.toString().trim() !== "") {
                  updatedFileId = currentFileId + "," + documentFileId;
                  Logger.log(
                    `[Register Template] Appending File ID: ${currentFileId} + ${documentFileId}`
                  );
                } else {
                  updatedFileId = documentFileId;
                  Logger.log(
                    `[Register Template] Setting first File ID: ${documentFileId}`
                  );
                }

                // Update the cell with the combined file IDs
                formSheet
                  .getRange(existingRowIndex + 1, fileIdColumnIndex + 1)
                  .setValue(updatedFileId);
                Logger.log(
                  `[Register Template] Updated File ID in spreadsheet: ${updatedFileId}`
                );
              }
            }
          }
        } catch (docError) {
          Logger.log(
            "[Register Template] Error creating document: " +
              docError.toString()
          );
          documentResult = {
            status: "error",
            message:
              "Template registered but document creation failed: " +
              docError.message,
          };
        }
      }
    } else {
      // No existing entry - create new with count 1
      count = 1;
      controlNumber = yearShort + "-" + ("0000" + count).slice(-4);
      Logger.log(
        "[Register Template] New template entry, control number: " +
          controlNumber
      );

      // Create an array with empty values for all columns
      // Make sure we have enough columns for File ID (which is after Queue)
      const maxColumnIndex = Math.max(
        queueColumnIndex + 1, // +1 for File ID column
        fieldsColumnIndex,
        headers.length - 1
      );
      const rowData = new Array(maxColumnIndex + 1).fill("");

      // Fill in the values we know
      rowData[0] = templateName;
      rowData[1] = currentYear;
      rowData[2] = count;
      rowData[3] = controlNumber;

      // Add fields in the right column
      if (fieldsColumnIndex !== -1 && fields) {
        rowData[fieldsColumnIndex] = fields;
      } else if (fields) {
        // Default fields column is index 4
        rowData[4] = fields;
      }

      // For new entries, just use the comma-separated email batch (no semicolons)
      rowData[queueColumnIndex] = emailBatch;

      // Create document if form data was provided
      if (formData && Array.isArray(formData) && formData.length > 0) {
        Logger.log(
          "[Register Template] Creating document with provided form data"
        );
        try {
          // Only pass control number to document creation if includeControlNumber is true
          const docControlNumber = includeControlNumber ? controlNumber : "";
          documentResult = createDoc(formData, docControlNumber, templateName);
          Logger.log(
            "[Register Template] Document created successfully: " +
              JSON.stringify(documentResult)
          );

          // Extract and store file ID
          if (documentResult && documentResult.status === "success") {
            // Extract file ID from documentUrl
            const docUrl = documentResult.documentUrl || "";
            if (docUrl) {
              documentFileId = docUrl.match(/\/d\/([^\/]+)/)?.[1] || "";
              Logger.log(
                `[Register Template] Extracted file ID: ${documentFileId}`
              );

              // Set File ID in rowData for the new entry
              if (documentFileId) {
                rowData[fileIdColumnIndex] = documentFileId;
                Logger.log(
                  `[Register Template] Added File ID to new row: ${documentFileId}`
                );
              }
            }
          }
        } catch (docError) {
          Logger.log(
            "[Register Template] Error creating document: " +
              docError.toString()
          );
          documentResult = {
            status: "error",
            message:
              "Template registered but document creation failed: " +
              docError.message,
          };
        }
      }

      Logger.log(
        "[Register Template] Appending new row data: " + JSON.stringify(rowData)
      );

      // Get the actual last row with data (to avoid skipping rows)
      const lastRow = formSheet.getLastRow();
      if (lastRow === 0) {
        // If there are no rows yet, add the data at row 1
        formSheet.getRange(1, 1, 1, rowData.length).setValues([rowData]);
      } else {
        // Add the data in the next row after the last row with data
        formSheet
          .getRange(lastRow + 1, 1, 1, rowData.length)
          .setValues([rowData]);
      }
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
      controlNumber: controlNumber,
      templateName: templateName,
      fileUrl: documentFileId
        ? `https://docs.google.com/document/d/${documentFileId}/edit`
        : `https://docs.google.com/document/d/${templateName}/edit`,
      queue: emailBatch, // Return just this batch of emails
      fields: fields,
      author: currentUserEmail,
      document: documentResult,
      fileId: documentFileId, // Include the file ID in the response
    };
  } catch (error) {
    Logger.log("[Register Template] ERROR: " + error.toString());
    throw error;
  } finally {
    lock.releaseLock();
  }
}
