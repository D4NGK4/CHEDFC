function createDoc(formData, controlNumber, templateName) {
  var isPrivate = false;

  Logger.log("[Document Creation] Starting process...");
  Logger.log(
    "[Document Creation] Received form data: " +
      JSON.stringify(formData, null, 2)
  );
  Logger.log("[Document Creation] Control Number: " + controlNumber);
  Logger.log("[Document Creation] Template Name: " + templateName);

  try {
    if (!templateName || typeof templateName !== "string") {
      throw new Error("Invalid template name provided");
    }

    // Get current user's email for TARGET PERSON field
    let currentUserEmail = "";
    try {
      currentUserEmail = Session.getActiveUser().getEmail();
      Logger.log("[Document Creation] Current user email: " + currentUserEmail);
    } catch (e) {
      Logger.log(
        "[Document Creation] Error getting user email: " + e.toString()
      );
    }

    const formDataObj = {};
    const formDataObjUpper = {};

    // Validate formData is an array
    if (!Array.isArray(formData)) {
      Logger.log(
        "[Document Creation] Form data is not an array, converting to expected format"
      );
      formData = [];
    }

    formData.forEach((item) => {
      // Skip items without a name property
      if (!item || !item.name) return;

      // If this is TARGET PERSON field, use current user's email
      if (item.name.toUpperCase() === "TARGET PERSON") {
        formDataObj[item.name] = currentUserEmail;
        formDataObjUpper[item.name.toUpperCase()] = currentUserEmail;
      } else {
        formDataObj[item.name] = item.value;
        formDataObjUpper[item.name.toUpperCase()] = item.value; // Store uppercase version for case-insensitive matching
      }
    });

    // Make sure TARGET PERSON is defined even if not in the form
    if (!formDataObj["TARGET PERSON"] && !formDataObjUpper["TARGET PERSON"]) {
      formDataObj["TARGET PERSON"] = currentUserEmail;
      formDataObjUpper["TARGET PERSON"] = currentUserEmail;
    }

    Logger.log(
      "[Document Creation] Form data object: " + JSON.stringify(formDataObj)
    );

    var data = [];

    // Only add control number to the data array if it exists
    if (controlNumber) {
      data.push(controlNumber);
    }

    const templateData = mapTemplateData(templateName);

    if (!templateData) {
      throw new Error("Template data not found for: " + templateName);
    }

    if (!templateData["Fields"]) {
      throw new Error("Template fields not defined for: " + templateName);
    }

    const fields = templateData["Fields"]
      .split(";")
      .map((f) => f.trim())
      .filter((f) => f);

    Logger.log(
      "[Document Creation] Template fields: " + JSON.stringify(fields)
    );

    // Ensure fields are properly formatted for the document generator
    fields.forEach((field) => {
      const [fieldName] = field.split(":").map((f) => f.trim());

      // Skip CONTROL NO field if we don't have a control number
      if (fieldName === "CONTROL NO" && !controlNumber) {
        Logger.log(
          "[Document Creation] Skipping CONTROL NO field as no control number was provided"
        );
        return;
      }

      if (fieldName !== "CONTROL NO") {
        // Special handling for TARGET PERSON field - use current user
        if (fieldName.toUpperCase() === "TARGET PERSON") {
          data.push(currentUserEmail || "");
        }
        // First try exact match
        else if (formDataObj[fieldName] !== undefined) {
          data.push(formDataObj[fieldName] || "");
        }
        // Then try case-insensitive match
        else if (formDataObjUpper[fieldName.toUpperCase()] !== undefined) {
          data.push(formDataObjUpper[fieldName.toUpperCase()] || "");
        }
        // Then try hyphenated field ID format (convert spaces to hyphens and lowercase)
        else {
          const fieldId = fieldName.toLowerCase().replace(/\s+/g, "-");
          if (formDataObj[fieldId] !== undefined) {
            data.push(formDataObj[fieldId] || "");
          } else {
            // If all else fails, push empty string
            data.push("");
            Logger.log(
              "[Document Creation] Field not found in form data: " + fieldName
            );
          }
        }
      }
    });

    Logger.log("[Document Creation] Final data array: " + JSON.stringify(data));

    try {
      // Ensure data is appropriate type for the document generator
      if (!Array.isArray(data)) {
        throw new Error("Data must be an array");
      }

      // Convert any complex values to strings to avoid type errors
      const safeData = data.map((item) => {
        if (item === null || item === undefined) return "";
        if (typeof item === "object") return JSON.stringify(item);
        return String(item);
      });

      Logger.log(
        "[Document Creation] Safe data array: " + JSON.stringify(safeData)
      );

      var doc = DocumentGenerator.createDocument(
        templateName,
        safeData,
        isPrivate
      );

      if (!doc) {
        throw new Error("Document creation returned no result");
      }
    } catch (docError) {
      Logger.log(
        "[Document Creation] Error in DocumentGenerator.createDocument: " +
          docError.toString()
      );
      throw new Error("Error creating document: " + docError.message);
    }

    Logger.log(
      "[Document Creation] Document created successfully: " +
        JSON.stringify(doc)
    );

    let fileId = extractFileId(doc);

    if (!fileId || fileId === "[object Object]") {
      Logger.log(
        "[Document Creation] WARNING: Could not extract proper file ID from document result"
      );
      if (typeof doc === "object") {
        for (const key in doc) {
          if (!doc.hasOwnProperty(key) || typeof doc[key] === "function")
            continue;

          const value = doc[key];
          if (typeof value === "string" && /^[a-zA-Z0-9_-]{20,}$/.test(value)) {
            fileId = value;
            Logger.log(
              "[Document Creation] Found potential file ID in property '" +
                key +
                "': " +
                fileId
            );
            break;
          }
        }
      }
    }

    if (!fileId) {
      throw new Error("Could not extract file ID from created document");
    }

    Logger.log("[Document Creation] Extracted file ID: " + fileId);

    return {
      status: "success",
      document: doc,
      documentUrl: `https://docs.google.com/document/d/${fileId}/edit`,
      controlNumber: controlNumber,
    };
  } catch (e) {
    Logger.log("[Document Creation] ERROR: " + e.toString());
    Logger.log("Stack: " + (e.stack || "No stack available"));

    return {
      status: "error",
      message: e.message || "Unknown error occurred",
      errorDetails: {
        name: e.name || "Error",
        stack: e.stack || "No stack trace available",
      },
    };
  }
}

function extractFileId(docUrl) {
  if (!docUrl) return null;

  try {
    // If docUrl is an object, try to extract URL property
    let url = docUrl;

    if (typeof docUrl === "object") {
      // Check for common URL properties
      if (docUrl.url) {
        url = docUrl.url;
      } else if (docUrl.documentUrl) {
        url = docUrl.documentUrl;
      } else if (docUrl.document) {
        url = docUrl.document;
      } else {
        // Convert the object to a string representation
        // But avoid using default toString() which produces [object Object]
        try {
          url = JSON.stringify(docUrl);
        } catch (e) {
          url = String(docUrl);
        }
      }
    }

    // Handle direct fileId return (not a URL)
    if (
      url &&
      typeof url === "string" &&
      !url.includes("/") &&
      !url.includes("?")
    ) {
      // If it looks like a file ID directly (no slashes or question marks), just return it
      return url;
    }

    // Handle various Google Doc URL formats
    if (typeof url === "string") {
      Logger.log("[Extract FileId] Processing URL: " + url);

      // Format 1: https://docs.google.com/document/d/{fileId}/edit
      let matches = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (matches && matches[1]) {
        return matches[1];
      }

      // Format 2: https://docs.google.com/open?id={fileId}
      matches = url.match(/[\?&]id=([a-zA-Z0-9_-]+)/);
      if (matches && matches[1]) {
        return matches[1];
      }

      // Format 3: Just directly extract any valid Google Doc ID pattern
      matches = url.match(/([a-zA-Z0-9_-]{25,})/);
      if (matches && matches[1]) {
        Logger.log(
          "[Extract FileId] Found ID pattern in string: " + matches[1]
        );
        return matches[1];
      }
    }
    Logger.log(
      "[Extract FileId] Could not extract file ID from: " +
        JSON.stringify(docUrl)
    );
    return null;
  } catch (e) {
    Logger.log("[Extract FileId] ERROR: " + e.toString());
    return null;
  }
}

function extractLatestStatus(statusArray) {
  try {
    if (!statusArray) {
      Logger.log("[Extract Status] No status array provided");
      return "Pending";
    }

    // If it's already a simple status string, return it
    if (typeof statusArray === "string") {
      return statusArray.trim() || "Pending";
    }

    // If it's not an array, try to convert it
    if (!Array.isArray(statusArray)) {
      Logger.log(
        "[Extract Status] Status is not an array, attempting to process:",
        JSON.stringify(statusArray)
      );

      // If it's an object with a status property
      if (
        statusArray &&
        typeof statusArray === "object" &&
        statusArray.status
      ) {
        return statusArray.status;
      }

      return "Pending";
    }

    Logger.log(
      "[Extract Status] Processing status array with " +
        statusArray.length +
        " entries"
    );

    const statusPriority = [
      "INITIAL STAMPED",
      "SIGNED",
      "COMPLETED",
      "APPROVED",
      "IN PROGRESS",
      "REQUESTING",
      "PENDING",
      "REJECTED",
    ];

    // Convert the array to a flat list of status strings
    const allStatuses = [];

    for (let i = 0; i < statusArray.length; i++) {
      const entry = statusArray[i];

      // Skip invalid entries
      if (!entry) continue;

      let status = "";

      // Handle different entry formats
      if (Array.isArray(entry)) {
        // Status is typically at index 2
        if (entry.length > 2 && typeof entry[2] === "string") {
          status = entry[2].trim();
        }
      } else if (typeof entry === "object" && entry.status) {
        status = entry.status;
      } else if (typeof entry === "string") {
        status = entry.trim();
      }

      if (status) {
        allStatuses.push(status);

        // If we find a completed status, return it immediately
        if (statusPriority.slice(0, 4).includes(status)) {
          Logger.log(
            "[Extract Status] Found completed status at index " +
              i +
              ": " +
              status
          );
          return status;
        }
      }
    }

    // If we have any statuses at all, return the highest priority one
    if (allStatuses.length > 0) {
      for (const priorityStatus of statusPriority) {
        if (allStatuses.includes(priorityStatus)) {
          Logger.log(
            "[Extract Status] Using priority status: " + priorityStatus
          );
          return priorityStatus;
        }
      }

      // No priority status found, return most recent
      const lastStatus = allStatuses[allStatuses.length - 1];
      Logger.log("[Extract Status] Using most recent status: " + lastStatus);
      return lastStatus;
    }

    // Default fallback
    Logger.log("[Extract Status] No valid status found, defaulting to Pending");
    return "Pending";
  } catch (e) {
    Logger.log("[Extract Status] ERROR: " + e.toString());
    return "Pending";
  }
}

function trackSubmission(fileId, fileName, author) {
  try {
    if (!fileId || fileId === "[object Object]") {
      Logger.log("[Track Submission] Invalid file ID: " + fileId);
      return false;
    }

    // Get the spreadsheet
    const ss = SpreadsheetApp.openById(
      "10kqLkE4WaPZGmpLExi6lO7lul-G8xJ6ONsl5C7k8aCE"
    );
    const sheet = ss.getSheets()[0]; // Assuming first sheet

    // Ensure we have the necessary columns
    ensureColumnsExist(sheet);

    // Get current user's email
    let currentUserEmail = "";
    try {
      currentUserEmail = Session.getActiveUser().getEmail();
      Logger.log("[Track Submission] Current user email: " + currentUserEmail);
    } catch (e) {
      Logger.log(
        "[Track Submission] Error getting user email: " + e.toString()
      );
    }

    // Get current status from the document generator
    let status = "Pending";
    try {
      const statusResult = DocumentGenerator.getRequestStatus(fileId);
      Logger.log(
        "[Track Submission] Raw status result:" +
          JSON.stringify(statusResult, null, 2)
      );
      status = extractLatestStatus(statusResult);
    } catch (statusErr) {
      Logger.log(
        "[Track Submission] ERROR getting status: " + statusErr.toString()
      );
    }

    Logger.log("[Track Submission] Final status: " + status);

    // Find next empty row or existing row with this fileId
    const lastRow = Math.max(1, sheet.getLastRow());
    let targetRow = lastRow + 1;

    // Check if this fileId already exists
    const data = sheet.getRange("A2:B" + lastRow).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][1] === fileId) {
        // File ID is in column B (index 1)
        targetRow = i + 2; // +2 because rows are 1-indexed and we have a header row
        break;
      }
    }

    // Create or update row with simplified data
    // Format: [fileName, fileId, email, status, lastUpdated, error]
    const rowData = [
      fileName,
      fileId,
      currentUserEmail,
      status,
      new Date(),
      "", // error message
    ];

    // Get the range size based on the number of columns we have or need
    const numColumns = Math.min(sheet.getLastColumn(), rowData.length);
    sheet
      .getRange(targetRow, 1, 1, numColumns)
      .setValues([rowData.slice(0, numColumns)]);

    Logger.log("[Track Submission] Successfully tracked submission: " + fileId);
    return true;
  } catch (e) {
    Logger.log("[Track Submission] ERROR: " + e.toString());
    return false;
  }
}

// Function to ensure all necessary columns exist in the sheet
function ensureColumnsExist(sheet) {
  try {
    // Define the expected headers
    const expectedHeaders = [
      "File Name",
      "File ID",
      "Email",
      "Status",
      "Last Updated",
      "Error Message",
    ];

    // Get existing headers
    const headerRow = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];

    // Check if we need to add new columns
    if (headerRow.length < expectedHeaders.length) {
      // Create a new header row with all expected columns
      sheet
        .getRange(1, 1, 1, expectedHeaders.length)
        .setValues([expectedHeaders]);
      Logger.log("[Ensure Columns] Added missing columns to sheet");
    }
  } catch (e) {
    Logger.log("[Ensure Columns] ERROR: " + e.toString());
  }
}

function getPersonalityDetails(personalityName) {
  try {
    const personalities = DocumentGenerator.getPersonalities();

    if (!personalities || !personalities.values) {
      throw new Error("No personalities data found");
    }

    // Log what we received to debug the structure
    Logger.log(
      "[Get Personality Details] First personality sample: " +
        JSON.stringify(personalities.values[0])
    );

    // Find the matching personality by name (index 1)
    const person = personalities.values.find((p) => p[1] === personalityName);

    if (!person) {
      Logger.log(
        "[Get Personality Details] No matching personality found for: " +
          personalityName
      );
      Logger.log(
        "[Get Personality Details] Available personalities: " +
          JSON.stringify(personalities.values.map((p) => p[1]))
      );
      return null;
    }

    // Log the found person to see its structure
    Logger.log(
      "[Get Personality Details] Found personality: " + JSON.stringify(person)
    );

    // Make sure we have all necessary data
    return {
      email: person[0] || "", // Email is at index 0
      name: person[1] || "", // Name is at index 1
      initials: person[2] || "", // Initials is at index 2
      level: person[4] || "", // Level is at index 4
    };
  } catch (e) {
    Logger.log("[Get Personality Details] ERROR: " + e.toString());
    return null;
  }
}

function getPersonalities() {
  try {
    const personalities = DocumentGenerator.getPersonalities();
    if (!personalities || !personalities.values) {
      throw new Error("No personalities data found");
    }

    // Extract just the names from the values array (name is in index 1)
    const names = personalities.values.map((person) => person[1]);

    return {
      status: "success",
      names: names,
      values: personalities.values, // Include the full data array
    };
  } catch (e) {
    Logger.log("[Get Personalities] ERROR: " + e.toString());
    return {
      status: "error",
      message: e.message || "Unknown error occurred",
    };
  }
}

function updateSubmissionStatus(fileId) {
  try {
    Logger.log("[Update Status] Starting update for: " + fileId);

    const statusResult = DocumentGenerator.getRequestStatus(fileId);

    // Add detailed logging
    Logger.log("[Update Status] Raw result type: " + typeof statusResult);
    Logger.log("Is array: " + Array.isArray(statusResult));

    let status = "Pending";

    if (Array.isArray(statusResult)) {
      Logger.log(
        "[Update Status] Result is array with length: " + statusResult.length
      );

      // Find the most recent completed status
      for (let i = statusResult.length - 1; i >= 0; i--) {
        const entry = statusResult[i];
        if (
          Array.isArray(entry) &&
          entry.length > 2 &&
          typeof entry[2] === "string"
        ) {
          const entryStatus = entry[2].trim();
          if (entryStatus) {
            Logger.log(
              `[Update Status] Found status at index ${i}: ${entryStatus}`
            );

            // If it's a completed status, use it immediately
            if (
              ["INITIAL STAMPED", "SIGNED", "COMPLETED", "APPROVED"].includes(
                entryStatus
              )
            ) {
              status = entryStatus;
              break;
            }

            // Otherwise keep the most recent non-empty status
            if (status === "Pending") {
              status = entryStatus;
            }
          }
        }
      }
    }

    Logger.log("[Update Status] Final status: " + status);

    const ss = SpreadsheetApp.openById(
      "10kqLkE4WaPZGmpLExi6lO7lul-G8xJ6ONsl5C7k8aCE"
    );
    const sheet = ss.getSheets()[0]; // Assuming first sheet

    // Find the row with this fileId
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange("A2:A" + lastRow).getValues();
    let targetRow = -1;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === fileId) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow === -1) {
      throw new Error("File ID not found in tracking sheet: " + fileId);
    }

    const rowData = sheet.getRange(targetRow, 1, 1, 4).getValues()[0];

    rowData[3] = status;

    sheet.getRange(targetRow, 1, 1, 4).setValues([rowData]);

    Logger.log(
      "[Update Status] Successfully updated status for: " +
        fileId +
        " with status: " +
        status
    );
    return true;
  } catch (e) {
    Logger.log("[Update Status] ERROR: " + e.toString());
    Logger.log("[Update Status] Error stack: " + e.stack);
    return false;
  }
}

function refreshAllStatuses() {
  try {
    const ss = SpreadsheetApp.openById(
      "10kqLkE4WaPZGmpLExi6lO7lul-G8xJ6ONsl5C7k8aCE"
    );
    const sheet = ss.getSheets()[0];

    // Get all data (including existing statuses)
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "success", updatedCount: 0 };

    // Get all data excluding the processed flag column (which has been removed)
    const dataRange = sheet.getRange("A2:D" + lastRow);
    const allData = dataRange.getValues();

    let updateCount = 0;
    const STATUS_PENDING = "Pending";
    const STATUS_PROCESSING = "Requesting initial";
    const STATUS_SIGNATURE_PROCESSING = "Requesting signature";
    const STATUS_SIGNATURE_STAMPED = "SIGNATURE STAMPED";

    // Update each row's status if needed
    for (let i = 0; i < allData.length; i++) {
      const row = allData[i];
      const fileId = row[1]; // FileID
      const currentStatus = row[3]; // Status

      if (!fileId) continue;

      // Skip rows that are in transition states
      if (
        currentStatus === STATUS_PENDING ||
        currentStatus === STATUS_PROCESSING ||
        currentStatus === STATUS_SIGNATURE_PROCESSING
      ) {
        continue;
      }

      // Skip rows that already completed with SIGNATURE_STAMPED
      if (currentStatus === STATUS_SIGNATURE_STAMPED) {
        continue;
      }

      // Use the same status retrieval logic as the queue system
      const statusResult = DocumentGenerator.getRequestStatus(fileId);
      const newStatus = getDocumentStatusFromGenerator(fileId);

      if (
        newStatus &&
        newStatus !== currentStatus &&
        isMoreAdvancedStatus(newStatus, currentStatus)
      ) {
        allData[i][3] = newStatus; // Update status
        updateCount++;
      }
    }

    if (updateCount > 0) {
      dataRange.setValues(allData);
    }

    Logger.log(`[Refresh Statuses] Updated ${updateCount} statuses`);
    return {
      status: "success",
      updatedCount: updateCount,
    };
  } catch (e) {
    Logger.log("[Refresh Statuses] ERROR: " + e.toString());
    return {
      status: "error",
      message: e.message || "Unknown error occurred",
    };
  }
}

function getDocumentStatusFromGenerator(fileId) {
  try {
    const results = DocumentGenerator.getRequestStatus(fileId);

    if (!Array.isArray(results) || results.length === 0) {
      return null;
    }

    if (hasStatus(results, "SIGNATURE STAMPED")) return "SIGNATURE STAMPED";
    if (hasStatus(results, "INITIAL STAMPED")) return "INITIAL STAMPED";

    let latestStatus = null;

    for (let i = results.length - 1; i >= 0; i--) {
      const item = results[i];
      if (Array.isArray(item) && item.length >= 3 && item[2] !== undefined) {
        latestStatus = item[2];
        break;
      }
    }

    return latestStatus;
  } catch (error) {
    Logger.log(
      `Error in getDocumentStatusFromGenerator for ${fileId}: ${error.toString()}`
    );
    return null;
  }
}

function hasStatus(results, statusToFind) {
  for (const item of results) {
    if (Array.isArray(item) && item.length >= 3 && item[2] === statusToFind) {
      return true;
    }
  }
  return false;
}

function isMoreAdvancedStatus(newStatus, currentStatus) {
  const statusHierarchy = [
    "Pending",
    "Requesting initial",
    "INITIAL STAMPED",
    "Requesting signature",
    "SIGNATURE STAMPED",
  ];

  const currentIndex = statusHierarchy.indexOf(currentStatus);
  const newIndex = statusHierarchy.indexOf(newStatus);

  if (currentIndex === -1) return true;
  if (newIndex === -1) return false;

  return newIndex >= currentIndex;
}

function getAllSubmissions() {
  try {
    const ss = SpreadsheetApp.openById(
      "10kqLkE4WaPZGmpLExi6lO7lul-G8xJ6ONsl5C7k8aCE"
    );
    const sheet = ss.getSheets()[0];

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return {
        status: "success",
        submissions: [],
      };
    }

    // Get all headers to check for personality columns
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];

    // Find column indexes
    const fileNameIndex = headers.indexOf("File Name");
    const fileIdIndex = headers.indexOf("File ID");
    const emailIndex = headers.indexOf("Email");
    const statusIndex = headers.indexOf("Status");
    const initialPersonalityIndex = headers.indexOf("Initial Personality");
    const signaturePersonalityIndex = headers.indexOf("Signature Personality");

    // Determine number of columns to fetch
    const numColumns = Math.max(
      fileNameIndex + 1,
      fileIdIndex + 1,
      emailIndex + 1,
      statusIndex + 1,
      initialPersonalityIndex + 1,
      signaturePersonalityIndex + 1
    );

    const dataRange = sheet.getRange(2, 1, lastRow - 1, numColumns);
    const submissionsData = dataRange.getValues();

    // Map to submission objects
    const submissions = submissionsData
      .map((row, index) => {
        const submission = {
          index: index + 1,
          fileName: row[fileNameIndex] || "",
          fileId: row[fileIdIndex] || "",
          email: row[emailIndex] || "",
          status: row[statusIndex] || "Pending",
          documentUrl: row[fileIdIndex]
            ? "https://docs.google.com/document/d/" + row[fileIdIndex] + "/edit"
            : "",
        };

        // Add personality information if available
        if (initialPersonalityIndex !== -1) {
          submission.initialPersonality = row[initialPersonalityIndex] || "";
        }
        if (signaturePersonalityIndex !== -1) {
          submission.signaturePersonality =
            row[signaturePersonalityIndex] || "";
        }

        return submission;
      })
      .filter((submission) => submission.fileId);

    Logger.log(
      "[Get Submissions] Retrieved " + submissions.length + " submissions"
    );

    return {
      status: "success",
      submissions: submissions,
    };
  } catch (e) {
    Logger.log("[Get Submissions] ERROR: " + e.toString());
    return {
      status: "error",
      message: e.message || "Unknown error occurred",
    };
  }
}
