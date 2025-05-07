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

    const formDataObj = {};
    formData.forEach((item) => {
      formDataObj[item.name] = item.value;
    });

    var data = [];
    if (controlNumber) {
      data.push(controlNumber);
    }

    const templateData = mapTemplateData(templateName);
    const fields = templateData["Fields"]
      .split(";")
      .map((f) => f.trim())
      .filter((f) => f);

    fields.forEach((field) => {
      const [fieldName] = field.split(":").map((f) => f.trim());
      if (fieldName !== "CONTROL NO") {
        data.push(formDataObj[fieldName] || "");
      }
    });

    Logger.log("[Document Creation] Final data array: " + JSON.stringify(data));

    var doc = DocumentGenerator.createDocument(templateName, data, isPrivate);

    if (!doc) {
      throw new Error("Document creation returned no result");
    }

    Logger.log(
      "[Document Creation] Document created successfully: " +
        JSON.stringify(doc)
    );

    let fileId = this.extractFileId(doc);

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

    // Get personality details from the enhanced modal
    let fileName = doc.name || templateName + " - " + controlNumber;

    // Get target person (author) from form data
    const targetPerson = formDataObj["TARGET PERSON"] || "";
    Logger.log("[Document Creation] Target Person: " + targetPerson);

    // Check for initial stamp requirements
    const needsInitialStamp =
      formDataObj.needsInitialStamp === true ||
      formDataObj.needsInitialStamp === "true";
    const initialPersonality = formDataObj.initialPersonality || targetPerson;

    // Check for signature stamp requirements
    const needsSignatureStamp =
      formDataObj.needsSignatureStamp === true ||
      formDataObj.needsSignatureStamp === "true";
    const signaturePersonality = formDataObj.signaturePersonality;

    Logger.log(
      "[Document Creation] Initial stamp needed: " +
        needsInitialStamp +
        ", personality: " +
        initialPersonality
    );
    Logger.log(
      "[Document Creation] Signature stamp needed: " +
        needsSignatureStamp +
        ", personality: " +
        signaturePersonality
    );

    // Track the submission with the new personality information
    const trackingResult = this.trackSubmissionWithPersonalities(
      fileId,
      fileName,
      needsInitialStamp ? initialPersonality : targetPerson,
      needsSignatureStamp ? signaturePersonality : "",
      needsInitialStamp || !!targetPerson,
      needsSignatureStamp
    );

    Logger.log(
      "[Document Creation] Tracked submission result: " + trackingResult
    );

    // If template has a level order, make sure the target person is first in queue
    if (templateData["Level Order"] && targetPerson) {
      try {
        // Set the target person as the first approver in the queue
        Logger.log(
          "[Document Creation] Setting target person as first approver in queue: " +
            targetPerson
        );
        DocumentGenerator.setFirstApprover(fileId, targetPerson);
      } catch (queueError) {
        Logger.log(
          "[Document Creation] ERROR setting approver queue: " +
            queueError.toString()
        );
      }
    }

    return {
      status: "success",
      document: doc,
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

function trackSubmissionWithPersonalities(
  fileId,
  fileName,
  initialPersonalityName,
  signaturePersonalityName,
  needsInitial,
  needsSignature
) {
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

    // Ensure we have the necessary columns - adjust if needed
    ensureColumnsExist(sheet);

    // Get initial personality details
    let initialPersonalityData = null;
    let initialPersonalityEmail = "";
    if (needsInitial && initialPersonalityName) {
      initialPersonalityData = this.getPersonalityDetails(
        initialPersonalityName
      );
      if (initialPersonalityData) {
        initialPersonalityEmail = initialPersonalityData.email || "";
        Logger.log(
          "[Track Submission] Initial personality email: " +
            initialPersonalityEmail
        );
      }
    }

    // Get signature personality details
    let signaturePersonalityData = null;
    let signaturePersonalityEmail = "";
    if (needsSignature && signaturePersonalityName) {
      signaturePersonalityData = this.getPersonalityDetails(
        signaturePersonalityName
      );
      if (signaturePersonalityData) {
        signaturePersonalityEmail = signaturePersonalityData.email || "";
        Logger.log(
          "[Track Submission] Signature personality email: " +
            signaturePersonalityEmail
        );
      }
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

    // Create or update row with new personality data
    // Format: [fileName, fileId, email(main), status, lastUpdated, error, needsInitial, initialPersonality, initialEmail, needsSignature, signaturePersonality, signatureEmail]
    const rowData = [
      fileName,
      fileId,
      initialPersonalityEmail || signaturePersonalityEmail, // Use initial email as main contact, fallback to signature
      status,
      new Date(),
      "", // error message
      needsInitial,
      initialPersonalityName || "",
      initialPersonalityEmail || "",
      needsSignature,
      signaturePersonalityName || "",
      signaturePersonalityEmail || "",
    ];

    // Get the range size based on the number of columns we have or need
    const numColumns = Math.min(sheet.getLastColumn(), rowData.length);
    sheet
      .getRange(targetRow, 1, 1, numColumns)
      .setValues([rowData.slice(0, numColumns)]);

    Logger.log(
      "[Track Submission] Successfully tracked submission with personalities: " +
        fileId
    );
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
      "Needs Initial",
      "Initial Personality",
      "Initial Email",
      "Needs Signature",
      "Signature Personality",
      "Signature Email",
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

// Modified for backward compatibility
function trackSubmission(fileId, fileName, personalityName) {
  // For backward compatibility, delegate to the new function
  return trackSubmissionWithPersonalities(
    fileId,
    fileName,
    personalityName,
    "",
    true,
    false
  );
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

function registerTemplate(templateName, levelOrder, targetPerson) {
  try {
    Logger.log(
      "[Register Template] Starting registration for: " + templateName
    );
    Logger.log("[Register Template] Level Order: " + levelOrder);
    Logger.log("[Register Template] Target Person: " + targetPerson);

    if (!templateName) {
      throw new Error("Template name is required");
    }

    if (!levelOrder) {
      throw new Error("Level order is required");
    }

    // Get existing templates
    const templates = DocumentGenerator.getTemplates();

    // Check if template exists
    if (!templates[templateName]) {
      throw new Error("Template not found: " + templateName);
    }

    // Store the target person and level order
    templates[templateName]["Level Order"] = levelOrder;
    templates[templateName]["Target Person"] = targetPerson || "";

    // Make sure the fields include TARGET PERSON:PERSONALITY if not already there
    let fields = templates[templateName]["Fields"] || "";
    if (!fields.includes("TARGET PERSON:PERSONALITY")) {
      // Check if fields ends with semicolon, if not add one
      if (fields && !fields.endsWith(";")) {
        fields += ";";
      }
      fields += " TARGET PERSON:PERSONALITY";
      templates[templateName]["Fields"] = fields;
      Logger.log(
        "[Register Template] Added TARGET PERSON field to template fields"
      );
    }

    // Save updated templates
    const result = DocumentGenerator.saveTemplates(templates);

    if (!result) {
      throw new Error("Failed to save template");
    }

    Logger.log(
      "[Register Template] Successfully registered template with level order"
    );
    return {
      status: "success",
      templateName: templateName,
      levelOrder: levelOrder,
      targetPerson: targetPerson || "Not explicitly provided",
      fields: templates[templateName]["Fields"],
    };
  } catch (e) {
    Logger.log("[Register Template] ERROR: " + e.toString());
    return {
      status: "error",
      message: e.message || "Unknown error occurred",
    };
  }
}
