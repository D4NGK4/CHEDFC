/**
 * Queue.js
 *
 * This script handles the signature request workflow by:
 * 1. Tracking document queues with multiple recipients
 * 2. Sending signature requests sequentially
 * 3. Tracking email status to prevent duplicate requests
 * 4. Cleaning up completed requests automatically
 */

//=============================================================================
// CONFIGURATION
//=============================================================================

const CONFIG = {
  // Spreadsheet where document queues are stored
  SPREADSHEET_ID: "1hm5BMdt7L9P_A9vZ5WGFYL_77KYnXBUTtZ8HihCfofk",
  SHEET_NAME: "Form Responses 1",

  // Column indices in spreadsheet (0-based)
  COLUMNS: {
    TEMPLATE_NAME: 0, // Column A: Template name
    YEAR: 1, // Column B: Year
    COUNT: 2, // Column C: Document counter
    CONTROL_NUMBER: 3, // Column D: Control number
    FIELDS: 4, // Column E: Form field mapping
    QUEUE: 5, // Column F: Email queue (semicolon separates batches, comma separates emails within a batch)
    FILE_ID: 6, // Column G: Google Doc file IDs (matching the batch structure in QUEUE)
    PROCESSED_QUEUE: 7, // Column H: Tracks emails that have already been sent requests
  },

  // Status values used for workflow tracking
  STATUS: {
    PENDING: "Pending",
    SIGNATURE_PROCESSING: "Requesting signature",
    SIGNATURE_STAMPED: "SIGNATURE STAMPED",
  },

  // List of statuses that indicate a completed signature
  COMPLETED_STATUSES: ["SIGNATURE STAMPED"],
};

//=============================================================================
// MAIN ENTRY POINT
//=============================================================================

/**
 * Main queue processing function - entry point for the script.
 * This function:
 * 1. Sets up necessary columns
 * 2. Gets all documents from queue
 * 3. Processes each document's signature requests sequentially
 */
function processQueue() {
  const startTime = new Date();
  Logger.log(`=== Starting queue processing at ${startTime.toISOString()} ===`);

  try {
    // Step 1: Ensure the processed queue column exists
    setupProcessedQueueColumn();

    // Step 2: Get all queue data from spreadsheet
    const queueData = getQueueData();
    if (!queueData || !queueData.values || queueData.values.length <= 1) {
      Logger.log("No documents in queue to process");
      return;
    }

    // Step 3: Process each row (each document template)
    for (let i = 1; i < queueData.values.length; i++) {
      const row = queueData.values[i];
      if (!row || row.length < 7) continue;

      const templateName = row[CONFIG.COLUMNS.TEMPLATE_NAME];
      const queueText = row[CONFIG.COLUMNS.QUEUE];
      const fileIdText = row[CONFIG.COLUMNS.FILE_ID];
      const rowIndex = i + 1; // 1-based row index for spreadsheet

      // Skip rows with no queue or file IDs
      if (!queueText || !fileIdText) continue;

      Logger.log(`Processing template ${templateName} at row ${rowIndex}`);

      // Step 4: Process this template's signature requests
      processTemplateSequentially(
        queueData.sheet,
        rowIndex,
        queueText,
        fileIdText
      );
    }
  } catch (error) {
    Logger.log(
      `CRITICAL ERROR in queue process: ${error.stack || error.toString()}`
    );
  } finally {
    const duration = (new Date() - startTime) / 1000;
    Logger.log(`=== Queue processing completed in ${duration} seconds ===`);
  }
}

//=============================================================================
// CORE PROCESSING FUNCTIONS
//=============================================================================

/**
 * Processes a template's signature requests in sequential order.
 *
 * The workflow:
 * 1. Loads current queues and already processed emails
 * 2. Organizes batches of emails and their file IDs
 * 3. For each batch, processes emails in order:
 *    - Checks if emails have completed signatures
 *    - Removes completed emails from queue
 *    - Sends requests to next pending email if needed
 *    - Tracks which emails have been processed to prevent duplicates
 * 4. Updates the spreadsheet with changes
 *
 * @param {Sheet} sheet - The spreadsheet sheet object
 * @param {Number} rowIndex - The 1-based row index in the spreadsheet
 * @param {String} queueText - Semicolon-separated batches of comma-separated emails
 * @param {String} fileIdText - Matching file IDs for each batch
 */
function processTemplateSequentially(sheet, rowIndex, queueText, fileIdText) {
  //--------------------------------------------------------------------------
  // STEP 1: Parse and organize queue data
  //--------------------------------------------------------------------------

  // Split all batches (batches are separated by semicolons)
  const emailBatches = queueText.split(";").map((batch) => batch.trim());
  let fileIdBatches = fileIdText.split(";").map((batch) => batch.trim());

  // Load the list of emails that have already been processed
  const processedQueueCell = sheet.getRange(
    rowIndex,
    CONFIG.COLUMNS.PROCESSED_QUEUE + 1
  );
  const processedQueueText = processedQueueCell.getValue() || "";
  const processedEmails = processedQueueText
    .split(",")
    .map((email) => email.trim())
    .filter((email) => email);

  Logger.log(
    `Processed emails from spreadsheet: ${processedEmails.join(", ")}`
  );

  //--------------------------------------------------------------------------
  // STEP 2: Handle special case - comma-separated file IDs
  //--------------------------------------------------------------------------

  // Sometimes file IDs might be comma-separated instead of matching the batch structure
  if (
    emailBatches.length > fileIdBatches.length &&
    fileIdBatches.length === 1
  ) {
    Logger.log(
      `Detected potential comma-separated file IDs instead of semicolon-separated batches`
    );

    // Check if the file IDs have commas and try to match them to email batches
    const fileIds = fileIdBatches[0].split(",").map((id) => id.trim());

    if (fileIds.length >= emailBatches.length) {
      // Reorganize file IDs to match the email batch structure
      fileIdBatches = [];
      for (let i = 0; i < emailBatches.length; i++) {
        if (i < fileIds.length) {
          fileIdBatches.push(fileIds[i]);
        } else {
          fileIdBatches.push(""); // Empty placeholder if no matching file ID
        }
      }

      Logger.log(
        `Restructured ${fileIds.length} file IDs into ${fileIdBatches.length} batches`
      );
    } else {
      Logger.log(
        `Cannot restructure file IDs: count mismatch (${fileIds.length} file IDs for ${emailBatches.length} email batches)`
      );
    }
  }

  // Verify batch counts match
  if (emailBatches.length !== fileIdBatches.length) {
    Logger.log(
      `ERROR: Mismatch between email batches (${emailBatches.length}) and file ID batches (${fileIdBatches.length})`
    );
    return;
  }

  //--------------------------------------------------------------------------
  // STEP 3: Process each batch in sequence
  //--------------------------------------------------------------------------

  // Tracking variables for queue state
  let processedFirstPending = false; // Tracks if we've processed one pending email this run
  let updatedQueue = false; // Tracks if we've made changes to the main queue
  let updatedProcessedQueue = false; // Tracks if we've made changes to the processed emails list
  let removedBatch = false; // Tracks if we've completely removed a batch
  let emailsInRemovedBatch = []; // Tracks emails in batches that were completely removed

  // Process each batch (each batch represents one file/document)
  for (let i = 0; i < emailBatches.length; i++) {
    const emailBatch = emailBatches[i];
    const fileIdBatch = fileIdBatches[i];

    // Skip empty batches
    if (!emailBatch || !fileIdBatch) {
      continue;
    }

    // Parse emails and file IDs from this batch
    const emails = emailBatch
      .split(",")
      .map((email) => email.trim())
      .filter((email) => !!email);

    const batchFileIds = fileIdBatch
      .split(",")
      .map((id) => id.trim())
      .filter((id) => !!id);

    // Skip invalid batches
    if (batchFileIds.length === 0 || emails.length === 0) {
      Logger.log(
        `Skipping batch ${i + 1} as it has no valid file IDs or emails`
      );
      continue;
    }

    // Get the file ID for this batch
    const fileId = batchFileIds[0];

    //--------------------------------------------------------------------------
    // STEP 4: Process emails in this batch sequentially
    //--------------------------------------------------------------------------

    let emailsToRemove = []; // List of indices of emails to remove
    let updatedEmailList = [...emails]; // Working copy of emails for this batch
    let foundPendingEmail = false; // Tracks if we found a pending email in this batch
    let currentBatchEmails = [...emails]; // Store all emails in this batch for later reference

    // Log batch information
    Logger.log(
      `Processing batch ${i + 1} with ${
        emails.length
      } emails for file ID ${fileId}`
    );

    // Process each email in this batch
    for (let j = 0; j < emails.length; j++) {
      const email = emails[j];

      // Check if this email has already been processed
      const isAlreadyProcessed = processedEmails.includes(email);
      if (isAlreadyProcessed) {
        Logger.log(
          `Email ${email} is already in processed list - skipping further checks`
        );
      }

      // Get detailed status from DocumentGenerator
      const results = DocumentGenerator.getRequestStatus(fileId);
      let isCompleted = false;
      let isInProgress = false;

      // Check all results to find status for this specific email
      if (Array.isArray(results)) {
        for (let k = 0; k < results.length; k++) {
          const item = results[k];
          if (Array.isArray(item) && item.length >= 3) {
            const resultEmail = String(item[0] || "").trim(); // Name field
            const resultEmailAddress = String(item[1] || "").trim(); // Email field
            const resultStatus = String(item[2] || "").trim(); // Status field

            // Log status details
            Logger.log(
              `Status check - Name: ${resultEmail}, Email: ${resultEmailAddress}, Status: ${resultStatus}`
            );

            // Check if this status applies to our target email
            const isEmailMatch =
              resultEmailAddress === email.trim() ||
              resultEmail === email.trim() ||
              (email.trim().includes("@") &&
                resultEmailAddress.includes("@") &&
                resultEmailAddress.trim() === email.trim());

            // Check for signature completion
            if (isEmailMatch && resultStatus === "SIGNATURE STAMPED") {
              isCompleted = true;
              Logger.log(
                `Found SIGNATURE STAMPED status for email ${email} (matched with ${resultEmailAddress})`
              );
              break;
            }

            // Check if signature request is in progress
            if (
              isEmailMatch &&
              (resultStatus.includes("signature") ||
                resultStatus.includes("SIGNATURE") ||
                resultStatus.includes("REQUEST") ||
                resultStatus.includes("request") ||
                resultStatus === "Requesting signature")
            ) {
              isInProgress = true;
              Logger.log(
                `Email ${email} has a signature request already in progress (${resultStatus})`
              );
            }
          }
        }
      }

      //--------------------------------------------------------------------------
      // STEP 5: Take action based on email status
      //--------------------------------------------------------------------------

      // Case 1: Email has completed signature - mark for removal
      if (isCompleted) {
        Logger.log(`Marking completed email for removal: ${email}`);
        emailsToRemove.push(j);
      }
      // Case 2: Email has request in progress or was previously processed
      else if (isInProgress || isAlreadyProcessed) {
        Logger.log(
          `Email ${email} already has a signature request in progress or was previously processed - NOT sending another email`
        );

        // Add to processed list if not already there
        if (!isAlreadyProcessed) {
          processedEmails.push(email);
          updatedProcessedQueue = true;
          Logger.log(`Added ${email} to processed emails list`);
        }

        // Skip further processing since an email is active
        if (!processedFirstPending) {
          processedFirstPending = true;
          Logger.log(
            `Active signature request found - will not send more requests this run`
          );
        }
      }
      // Case 3: Need to check status through helper function
      else {
        // Get status through helper function
        const emailStatus = getDocumentStatus(fileId, email);

        Logger.log(
          `Batch ${i + 1}, Email ${j + 1}/${
            emails.length
          }: ${email}, Status: ${emailStatus}`
        );

        // Case 3a: Email is completed
        if (isCompletedStatus(emailStatus)) {
          Logger.log(
            `Marking completed email for removal via normal check: ${email}`
          );
          emailsToRemove.push(j);
        }
        // Case 3b: Email is pending and we haven't processed another email yet
        else if (
          emailStatus === CONFIG.STATUS.PENDING &&
          !foundPendingEmail &&
          !processedFirstPending &&
          !isAlreadyProcessed
        ) {
          // Found first pending email to process
          foundPendingEmail = true;

          // Send the signature request
          Logger.log(`Requesting signature for first pending email: ${email}`);
          requestStampSignature(email, fileId);

          // Mark as processed to prevent duplicate requests
          processedEmails.push(email);
          updatedProcessedQueue = true;
          Logger.log(`Added ${email} to processed emails list`);

          processedFirstPending = true;
        }
        // Case 3c: Email is in signature process
        else if (emailStatus === CONFIG.STATUS.SIGNATURE_PROCESSING) {
          Logger.log(
            `Email ${email} already has a signature request in progress - skipping`
          );

          // Add to processed list if not already there
          if (!isAlreadyProcessed) {
            processedEmails.push(email);
            updatedProcessedQueue = true;
            Logger.log(`Added ${email} to processed emails list`);
          }

          // Skip further processing
          if (!processedFirstPending) {
            processedFirstPending = true;
            Logger.log(
              `Active signature request found - will not send more requests this run`
            );
          }
        }
      }
    }

    //--------------------------------------------------------------------------
    // STEP 6: Update batch by removing completed emails
    //--------------------------------------------------------------------------

    if (emailsToRemove.length > 0) {
      Logger.log(
        `Removing ${emailsToRemove.length} completed emails from batch ${i + 1}`
      );

      // Remove emails in reverse order to maintain correct indices
      for (let j = emailsToRemove.length - 1; j >= 0; j--) {
        const indexToRemove = emailsToRemove[j];
        updatedEmailList.splice(indexToRemove, 1);
      }

      // Update batch with remaining emails or remove batch if empty
      if (updatedEmailList.length > 0) {
        emailBatches[i] = updatedEmailList.join(",");
      } else {
        // If this batch is completely done, remember its emails for removal from processed queue
        Logger.log(`Batch ${i + 1} completed - all emails have been signed`);
        emailsInRemovedBatch = emailsInRemovedBatch.concat(currentBatchEmails);

        // Remove entire batch if no emails remain
        emailBatches.splice(i, 1);
        fileIdBatches.splice(i, 1);
        removedBatch = true;
        i--; // Adjust index since we removed an item
      }

      updatedQueue = true;
    }

    // Stop after processing one batch with either a pending email or active request
    if (foundPendingEmail || processedFirstPending) {
      break;
    }
  }

  //--------------------------------------------------------------------------
  // STEP 7: Clean up processed queue for completed batches
  //--------------------------------------------------------------------------

  // If we removed at least one complete batch, remove its emails from processed queue
  if (removedBatch && emailsInRemovedBatch.length > 0) {
    Logger.log(
      `Removing ${emailsInRemovedBatch.length} emails from processed queue because their batches are complete`
    );

    // Filter out emails from completed batches
    const updatedProcessedEmails = processedEmails.filter(
      (email) => !emailsInRemovedBatch.includes(email)
    );

    // Only update if there was an actual change
    if (updatedProcessedEmails.length !== processedEmails.length) {
      processedEmails.length = 0; // Clear array
      processedEmails.push(...updatedProcessedEmails); // Add filtered emails
      updatedProcessedQueue = true;
      Logger.log(
        `Updated processed emails list: ${processedEmails.join(", ")}`
      );
    }
  }

  //--------------------------------------------------------------------------
  // STEP 8: Update spreadsheet with changes
  //--------------------------------------------------------------------------

  if (updatedQueue || updatedProcessedQueue) {
    updateSpreadsheetRow(
      sheet,
      rowIndex,
      emailBatches,
      fileIdBatches,
      processedEmails
    );
    Logger.log(
      `Updated queue after processing - will process next email in next run`
    );
    return;
  }

  // If all batches are empty, clear the row completely
  if (emailBatches.length === 0) {
    Logger.log(`Clearing row ${rowIndex} as all batches are completed`);
    sheet.getRange(rowIndex, CONFIG.COLUMNS.QUEUE + 1).setValue("");
    sheet.getRange(rowIndex, CONFIG.COLUMNS.FILE_ID + 1).setValue("");
    sheet.getRange(rowIndex, CONFIG.COLUMNS.PROCESSED_QUEUE + 1).setValue("");
  }
}

//=============================================================================
// HELPER FUNCTIONS - SPREADSHEET OPERATIONS
//=============================================================================

/**
 * Gets all queue data from the spreadsheet.
 *
 * @return {Object} Object containing sheet, values array, and headers
 */
function getQueueData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet) {
      throw new Error(`Sheet '${CONFIG.SHEET_NAME}' not found`);
    }

    // Get all data from the sheet
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    return {
      sheet: sheet,
      values: values,
      headers: values[0],
    };
  } catch (error) {
    Logger.log(`Error getting queue data: ${error.toString()}`);
    throw error;
  }
}

/**
 * Updates a spreadsheet row with modified queue data.
 *
 * @param {Sheet} sheet - The spreadsheet sheet object
 * @param {Number} rowIndex - The 1-based row index in the spreadsheet
 * @param {Array} emailBatches - Array of email batches
 * @param {Array} fileIdBatches - Array of file ID batches
 * @param {Array} processedEmails - Array of emails that have been processed
 */
function updateSpreadsheetRow(
  sheet,
  rowIndex,
  emailBatches,
  fileIdBatches,
  processedEmails = []
) {
  // Prepare updated values
  const updatedQueueText =
    emailBatches.length > 0 ? emailBatches.join(";") : "";
  const updatedFileIdText =
    fileIdBatches.length > 0 ? fileIdBatches.join(";") : "";
  const updatedProcessedQueueText =
    processedEmails.length > 0 ? processedEmails.join(",") : "";

  // Update the queue and file IDs columns
  sheet.getRange(rowIndex, CONFIG.COLUMNS.QUEUE + 1).setValue(updatedQueueText);
  sheet
    .getRange(rowIndex, CONFIG.COLUMNS.FILE_ID + 1)
    .setValue(updatedFileIdText);

  // Update the processed queue column
  sheet
    .getRange(rowIndex, CONFIG.COLUMNS.PROCESSED_QUEUE + 1)
    .setValue(updatedProcessedQueueText);

  Logger.log(
    `Updated row ${rowIndex} with ${emailBatches.length} remaining batches. Processed emails: ${updatedProcessedQueueText}`
  );
}

/**
 * Sets up the PROCESSED_QUEUE column if it doesn't exist.
 * This ensures the script can track which emails have been processed.
 */
function setupProcessedQueueColumn() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet) {
      Logger.log(`Sheet '${CONFIG.SHEET_NAME}' not found`);
      return;
    }

    // Get header row
    const headerRow = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];

    // Check if PROCESSED_QUEUE column exists
    let processedQueueColumnIndex = -1;
    for (let i = 0; i < headerRow.length; i++) {
      if (headerRow[i] === "Processed Queue") {
        processedQueueColumnIndex = i;
        break;
      }
    }

    // If column doesn't exist, add it
    if (processedQueueColumnIndex === -1) {
      const newColumnIndex = CONFIG.COLUMNS.PROCESSED_QUEUE + 1;

      // Add header if the column doesn't exist
      if (headerRow.length < newColumnIndex) {
        sheet.getRange(1, newColumnIndex).setValue("Processed Queue");
        Logger.log(`Added 'Processed Queue' column at index ${newColumnIndex}`);
      }
    }
  } catch (error) {
    Logger.log(`Error setting up processed queue column: ${error.toString()}`);
  }
}

//=============================================================================
// HELPER FUNCTIONS - STATUS CHECKS AND SIGNATURE REQUESTS
//=============================================================================

/**
 * Checks if a status is considered "completed".
 *
 * @param {String} status - The status to check
 * @return {Boolean} True if the status indicates completion
 */
function isCompletedStatus(status) {
  return CONFIG.COMPLETED_STATUSES.includes(status);
}

/**
 * Checks if a specific status exists in the results array.
 *
 * @param {Array} results - Array of status results
 * @param {String} targetStatus - The status to search for
 * @return {Boolean} True if the status was found
 */
function hasStatus(results, targetStatus) {
  if (!Array.isArray(results)) return false;

  for (let i = 0; i < results.length; i++) {
    const item = results[i];
    if (Array.isArray(item) && item.length >= 3) {
      const status = item[2];
      if (status === targetStatus) return true;
    }
  }

  return false;
}

/**
 * Gets document status from DocumentGenerator for a specific email and file ID.
 *
 * @param {String} fileId - The Google Doc file ID
 * @param {String} email - The email address to check status for
 * @return {String} The status of the document for this email
 */
function getDocumentStatus(fileId, email) {
  try {
    if (!fileId || typeof fileId !== "string") {
      return CONFIG.STATUS.PENDING; // Default to Pending for invalid IDs
    }

    // Use DocumentGenerator to get status
    const results = DocumentGenerator.getRequestStatus(fileId);

    if (!Array.isArray(results) || results.length === 0) {
      return CONFIG.STATUS.PENDING; // Not sent yet
    }

    // If email is specified, look for status specifically for this email
    if (email) {
      Logger.log(
        `Looking for status specific to email: ${email} for file ${fileId}`
      );

      // Look for entries matching this email
      for (let i = 0; i < results.length; i++) {
        const item = results[i];
        if (Array.isArray(item) && item.length >= 3) {
          const resultName = String(item[0] || "").trim();
          const resultEmail = String(item[1] || "").trim();
          const status = String(item[2] || "").trim();

          // Check if either the name or email matches
          const isMatch =
            resultEmail === email.trim() ||
            resultName === email.trim() ||
            (email.includes("@") &&
              resultEmail.includes("@") &&
              resultEmail === email);

          // If this result is for our target email
          if (isMatch) {
            Logger.log(
              `Found status for ${email} (matched with ${resultEmail}): ${status}`
            );

            // Check status hierarchy
            if (status === "SIGNATURE STAMPED") {
              return CONFIG.STATUS.SIGNATURE_STAMPED;
            } else if (
              status &&
              (status.includes("signature") ||
                status.includes("SIGNATURE") ||
                status.includes("REQUEST") ||
                status.includes("request"))
            ) {
              return CONFIG.STATUS.SIGNATURE_PROCESSING;
            }

            // Default to processing if we found the email but status isn't clear
            return CONFIG.STATUS.SIGNATURE_PROCESSING;
          }
        }
      }

      // If we didn't find any status for this email, it's still pending
      return CONFIG.STATUS.PENDING;
    }

    // If no email specified, use the original logic (for backward compatibility)
    // Check for signature stamped status (highest priority)
    if (hasStatus(results, "SIGNATURE STAMPED")) {
      return CONFIG.STATUS.SIGNATURE_STAMPED;
    }

    // Any other status means the document is in the signature process
    return CONFIG.STATUS.SIGNATURE_PROCESSING;
  } catch (error) {
    Logger.log(`Error getting status for ${fileId}: ${error.toString()}`);
    return CONFIG.STATUS.PENDING; // Default to Pending on error
  }
}

/**
 * Requests a signature stamp for a given email and file ID.
 * This function:
 * 1. Validates the email and file ID
 * 2. Checks if a request has already been sent
 * 3. Sends the signature request if needed
 *
 * @param {Array|String} email - Email address or array with one email
 * @param {String} fileId - Google Doc file ID to request signature for
 */
function requestStampSignature(email, fileId) {
  try {
    // Handle both string and array inputs
    const emailAddress = Array.isArray(email) ? email[0] : email;

    // Validate inputs
    if (!emailAddress || !fileId || typeof fileId !== "string") {
      Logger.log(`Invalid request: missing email or fileId`);
      return;
    }

    if (!emailAddress.includes("@")) {
      Logger.log(`Invalid email address: ${emailAddress}`);
      return;
    }

    // Check if we've already sent a request to this email
    const results = DocumentGenerator.getRequestStatus(fileId);
    let alreadyRequested = false;

    if (Array.isArray(results)) {
      for (let i = 0; i < results.length; i++) {
        const item = results[i];
        if (Array.isArray(item) && item.length >= 3) {
          const resultEmail = String(item[0] || "").trim();
          const resultEmailAddress = String(item[1] || "").trim();
          const resultStatus = String(item[2] || "").trim();

          // Check if this email already has a request
          const isEmailMatch =
            resultEmailAddress === emailAddress.trim() ||
            resultEmail === emailAddress.trim() ||
            (resultEmailAddress.includes("@") &&
              emailAddress.includes("@") &&
              resultEmailAddress.trim() === emailAddress.trim());

          if (
            isEmailMatch &&
            (resultStatus.includes("signature") ||
              resultStatus.includes("SIGNATURE") ||
              resultStatus.includes("REQUEST") ||
              resultStatus.includes("request"))
          ) {
            alreadyRequested = true;
            Logger.log(
              `Email ${emailAddress} already has a request (${resultStatus}) - skipping`
            );
            break;
          }
        }
      }
    }

    // Only send if no previous request exists
    if (!alreadyRequested) {
      Logger.log(`Requesting stamp signature for ${fileId} to ${emailAddress}`);

      try {
        DocumentGenerator.requestStampSignature(emailAddress, fileId);
        Logger.log(
          `Successfully requested signature from ${emailAddress} for ${fileId}`
        );
      } catch (error) {
        Logger.log(
          `Failed to request signature from ${emailAddress}: ${error.toString()}`
        );
      }
    } else {
      Logger.log(`Skipped duplicate signature request to ${emailAddress}`);
    }
  } catch (error) {
    Logger.log(`Error in requestStampSignature: ${error.toString()}`);
  }
}

//=============================================================================
// TEST FUNCTIONS
//=============================================================================

/**
 * Test function for debugging the queue process.
 */
function testQueue() {
  processQueue();
}

/**
 * Refreshes all document queue statuses.
 * Useful for manually updating all document statuses.
 *
 * @return {Object} Status object with success/error details
 */
function refreshAllStatuses() {
  try {
    const queueData = getQueueData();
    if (!queueData || !queueData.values || queueData.values.length <= 1) {
      return { status: "success", message: "No documents to refresh" };
    }

    let updateCount = 0;

    // Process each template row
    for (let i = 1; i < queueData.values.length; i++) {
      const row = queueData.values[i];
      if (!row || row.length < 7) continue;

      const queueText = row[CONFIG.COLUMNS.QUEUE];
      const fileIdText = row[CONFIG.COLUMNS.FILE_ID];
      const rowIndex = i + 1;

      // Skip if no queue or file IDs
      if (!queueText || !fileIdText) continue;

      processTemplateSequentially(
        queueData.sheet,
        rowIndex,
        queueText,
        fileIdText
      );
      updateCount++;
    }

    return {
      status: "success",
      updatedCount: updateCount,
      message: `Refreshed ${updateCount} template queues`,
    };
  } catch (error) {
    Logger.log(`Error refreshing statuses: ${error.toString()}`);
    return {
      status: "error",
      message: error.message || "Unknown error",
    };
  }
}
