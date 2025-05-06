const CONFIG = {
  SSID: "10kqLkE4WaPZGmpLExi6lO7lul-G8xJ6ONsl5C7k8aCE",
  SHEETNAME: "Sheet1",
  STATUS: {
    PENDING: "Pending",
    PROCESSING: "Requesting initial",
    INITIAL_STAMPED: "INITIAL STAMPED",
    SIGNATURE_PROCESSING: "Requesting signature",
    SIGNATURE_STAMPED: "SIGNATURE STAMPED",
    ERROR: "Error"
  },
  COLUMNS: {
    FILE_NAME: 0,
    FILE_ID: 1,
    EMAIL: 2,
    STATUS: 3,
    LAST_UPDATED: 4,
    ERROR_MESSAGE: 5,
    NEEDS_INITIAL: 6,
    INITIAL_PERSONALITY: 7,
    INITIAL_EMAIL: 8,
    NEEDS_SIGNATURE: 9,
    SIGNATURE_PERSONALITY: 10,
    SIGNATURE_EMAIL: 11
  },
};

class DocumentProcessingError extends Error {
  constructor(message, documentId = null) {
    super(message);
    this.name = "DocumentProcessingError";
    this.documentId = documentId;
  }
}

class DocumentNotFoundError extends DocumentProcessingError {
  constructor(documentId) {
    super(`Document not found: ${documentId}`, documentId);
    this.name = "DocumentNotFoundError";
  }
}

class InvalidStateTransitionError extends DocumentProcessingError {
  constructor(currentStatus, newStatus, documentId) {
    super(`Document ${documentId} is still in ${currentStatus} state`, documentId);
    this.name = "StatusNotification";
    this.currentStatus = currentStatus;
    this.newStatus = newStatus;
  }
}

/**
 * Main queue function 
 */
function queue() {
  const startTime = new Date();
  Logger.log(`=== Starting queue processing at ${startTime.toISOString()} ===`);
  
  try {    
    // 1. First update all document statuses
    updateAllDocumentStatuses();
    
    // 2. Process any documents that need signature requests
    processInitialStampedDocuments();
    
    // 3. Process pending documents (staggered processing)
    processPendingDocuments();
    
  } catch (error) {
    Logger.log(`CRITICAL ERROR in queue process: ${error.stack || error.toString()}`);
    notifyAdmin(`Queue processing failed: ${error.message}`);
  } finally {
    const duration = (new Date() - startTime) / 1000;
    Logger.log(`=== Queue processing completed in ${duration} seconds ===`);
  }
}

/**
 * Updates all document statuses from DocumentGenerator
 */
function updateAllDocumentStatuses() {
  const { sheet, values } = getSheetData();
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row || row.length < 4) continue;
    
    const fileId = row[CONFIG.COLUMNS.FILE_ID];
    if (!fileId) continue;
    
    try {
      updateDocumentStatus(sheet, i + 1, row);
    } catch (error) {
      // Only handle actual errors, not status notifications
      if (!(error instanceof InvalidStateTransitionError)) {
        handleDocumentError(sheet, i + 1, row, error);
      }
    }
  }
}

function updateDocumentStatus(sheet, rowNum, row) {
  const fileId = row[CONFIG.COLUMNS.FILE_ID];
  const currentStatus = row[CONFIG.COLUMNS.STATUS];
  
  // Skip if already completed
  if (currentStatus === CONFIG.STATUS.SIGNATURE_STAMPED) return;
  
  // Get current status from DocumentGenerator
  const docStatus = getDocumentStatusFromGenerator(fileId);
  if (!docStatus) {
    Logger.log(`No status returned from DocumentGenerator for ${fileId}, keeping current status: ${currentStatus}`);
    return;
  }
  
  // Update status if changed
  if (docStatus !== currentStatus) {
    Logger.log(`Status transition for ${fileId}: ${currentStatus} → ${docStatus}`);
    
    const updates = {
      status: docStatus,
      lastUpdated: new Date()
    };
    
    updateRow(sheet, rowNum, updates);
  }
}

/**
 * Processes pending documents for initial stamping
 */
function processPendingDocuments() {
  const { sheet, values } = getSheetData();
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row || row.length < 4) continue;
    
    const fileId = row[CONFIG.COLUMNS.FILE_ID];
    const email = row[CONFIG.COLUMNS.EMAIL];
    const status = row[CONFIG.COLUMNS.STATUS];
    const needsInitial = row[CONFIG.COLUMNS.NEEDS_INITIAL];
    const initialEmail = row[CONFIG.COLUMNS.INITIAL_EMAIL];
    
    // Only process documents that are in PENDING state and haven't been marked as ERROR
    if (!fileId || status !== CONFIG.STATUS.PENDING) continue;
    
    try {
      // Only process documents that need initial stamps and have initial emails
      if (needsInitial && initialEmail) {
        processDocumentForInitial(sheet, i + 1, row);
        break; // Process one at a time to avoid race conditions
      }
      // If document doesn't need an initial stamp but needs signature
      else if (!needsInitial && row[CONFIG.COLUMNS.NEEDS_SIGNATURE] && row[CONFIG.COLUMNS.SIGNATURE_EMAIL]) {
        // Skip directly to signature processing
        processDocumentForSignature(sheet, i + 1, row);
        break;
      }
    } catch (error) {
      handleDocumentError(sheet, i + 1, row, error);
    }
  }
}

function processDocumentForInitial(sheet, rowNum, row) {
  const fileId = row[CONFIG.COLUMNS.FILE_ID];
  const fileName = row[CONFIG.COLUMNS.FILE_NAME];
  const initialEmail = row[CONFIG.COLUMNS.INITIAL_EMAIL];
  
  Logger.log(`Processing document for initial: ${fileName}, File ID: ${fileId}, Email: ${initialEmail}`);

  updateRow(sheet, rowNum, {
    status: CONFIG.STATUS.PROCESSING,
    lastUpdated: new Date()
  });
  
  // Request initial stamp 
  const result = requestInitialForDocument(fileId, initialEmail);
  Logger.log(`Initial request result for ${fileId}: ${JSON.stringify(result)}`);
}

/**
 * Processes documents that have been initially stamped
 */
function processInitialStampedDocuments() {
  const { sheet, values } = getSheetData();
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row || row.length < 4) continue;
    
    const fileId = row[CONFIG.COLUMNS.FILE_ID];
    const status = row[CONFIG.COLUMNS.STATUS];
    const needsSignature = row[CONFIG.COLUMNS.NEEDS_SIGNATURE];
    const signatureEmail = row[CONFIG.COLUMNS.SIGNATURE_EMAIL];
    
    if (!fileId || status !== CONFIG.STATUS.INITIAL_STAMPED) continue;
    
    try {
      // Only process documents that need signature stamps and have signature emails
      if (needsSignature && signatureEmail) {
        processDocumentForSignature(sheet, i + 1, row);
        break; // Process one at a time to avoid race conditions
      }
    } catch (error) {
      handleDocumentError(sheet, i + 1, row, error);
    }
  }
}

function processDocumentForSignature(sheet, rowNum, row) {
  const fileId = row[CONFIG.COLUMNS.FILE_ID];
  const fileName = row[CONFIG.COLUMNS.FILE_NAME];
  const signatureEmail = row[CONFIG.COLUMNS.SIGNATURE_EMAIL];
  
  Logger.log(`Processing document for signature: ${fileName}, File ID: ${fileId}, Email: ${signatureEmail}`);
  
  // Update status
  updateRow(sheet, rowNum, {
    status: CONFIG.STATUS.SIGNATURE_PROCESSING,
    lastUpdated: new Date()
  });
  
  // Request signature with the specific signature email
  const result = sendForSignature(fileId, signatureEmail);
  Logger.log(`Signature request result for ${fileId}: ${JSON.stringify(result)}`);
}

/**
 * Recovers stuck documents
 */


// ======================
// Core Service Functions
// ======================

function requestInitialForDocument(fileId, email) {
  try {
    validateDocumentId(fileId);
    validateEmail(email);
    
    const results = DocumentGenerator.requestStampInitial(email, fileId);
    if (!results || results.error) {
      throw new DocumentProcessingError(results?.error || "Unknown error from DocumentGenerator", fileId);
    }
    return results;
  } catch (error) {
    throw new DocumentProcessingError(`Failed to request initial for ${fileId}: ${error.message}`, fileId);
  }
}

function sendForSignature(fileId, email) {
  try {
    validateDocumentId(fileId);
    validateEmail(email);
    
    const results = DocumentGenerator.requestStampSignature(email, fileId);
    if (!results || results.error) {
      throw new DocumentProcessingError(results?.error || "Unknown error from DocumentGenerator", fileId);
    }
    return results;
  } catch (error) {
    throw new DocumentProcessingError(`Failed to request signature for ${fileId}: ${error.message}`, fileId);
  }
}

function getDocumentStatusFromGenerator(fileId) {
  try {
    validateDocumentId(fileId);
    
    const results = DocumentGenerator.getRequestStatus(fileId);
    if (!Array.isArray(results) || results.length === 0) {
      throw new DocumentNotFoundError(fileId);
    }
    
    let mostRecentStatus = null;
    let mostRecentTimestamp = 0;
    
    for (const item of results) {
      if (Array.isArray(item) && item.length >= 4) {
        const status = item[2];
        const timestampStr = item[3];
        
        try {
          const timestamp = new Date(timestampStr).getTime();
          if (timestamp > mostRecentTimestamp) {
            mostRecentTimestamp = timestamp;
            mostRecentStatus = status;
          }
        } catch (e) {
          Logger.log(`Invalid timestamp format: ${timestampStr}`);
        }
      }
    }
    
    if (!mostRecentStatus) {
      throw new DocumentProcessingError("No valid status found in response", fileId);
    }
    
    return mostRecentStatus;
  } catch (error) {
    throw new DocumentProcessingError(`Failed to get status for ${fileId}: ${error.message}`, fileId);
  }
}

// ======================
// Helper Functions
// ======================

function getSheetData() {
  const sheet = SpreadsheetApp.openById(CONFIG.SSID).getSheetByName(CONFIG.SHEETNAME);
  const dataRange = sheet.getDataRange();
  return {
    sheet,
    values: dataRange.getValues(),
    range: dataRange
  };
}

function updateRow(sheet, rowNum, updates) {
  const range = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn());
  const row = range.getValues()[0];
  
  for (const [key, value] of Object.entries(updates)) {
    const colIndex = CONFIG.COLUMNS[key.toUpperCase()];
    if (colIndex !== undefined && colIndex < row.length) {
      row[colIndex] = value;
    }
  }
  
  range.setValues([row]);
  SpreadsheetApp.flush();
}

function handleDocumentError(sheet, rowNum, row, error) {
  const fileId = row[CONFIG.COLUMNS.FILE_ID];
  
  Logger.log(`Error processing document ${fileId}: ${error.stack || error.toString()}`);
  
  updateRow(sheet, rowNum, {
    lastUpdated: new Date(),
    errorMessage: error.message
  });
}

function markAsFailed(sheet, rowNum, row, errorMessage) {
  updateRow(sheet, rowNum, {
    status: CONFIG.STATUS.ERROR,
    lastUpdated: new Date(),
    errorMessage: errorMessage
  });
}

function validateDocumentId(fileId) {
  if (!fileId || typeof fileId !== 'string') {
    throw new DocumentProcessingError("Invalid document ID");
  }
}

function validateEmail(email) {
  if (!email || typeof email !== 'string' || !email.includes('@')) {
    throw new DocumentProcessingError("Invalid email address");
  }
}

function notifyAdmin(message) {
  // Implement your notification logic (email, Slack, etc.)
  Logger.log(`ADMIN NOTIFICATION: ${message}`);
}

// ======================
// Utility Functions

// ======================

function addToQueue(fileName, fileId, email) {
  try {
    validateDocumentId(fileId);
    validateEmail(email);
    
    const { sheet } = getSheetData();
    sheet.appendRow([
      fileName,
      fileId,
      email,
      CONFIG.STATUS.PENDING,
      new Date(), // last updated
      "", // error message
      false, // needs initial
      "", // initial personality
      "", // initial email
      false, // needs signature
      "", // signature personality
      "" // signature email
    ]);
    Logger.log(`Added document to queue: ${fileName}, File ID: ${fileId}`);
  } catch (error) {
    Logger.log(`Error adding document to queue: ${error.stack || error.toString()}`);
    throw error;
  }
}

function resetDocumentData() {
  const { sheet, values } = getSheetData();
  
  for (let i = 1; i < values.length; i++) {
    updateRow(sheet, i + 1, {
      errorMessage: "",
      needsInitial: false,
      initialPersonality: "",
      initialEmail: "",
      needsSignature: false,
      signaturePersonality: "",
      signatureEmail: ""
    });
  }
  Logger.log("Reset all document data and error states");
}

function setupSheetColumns() {
  const { sheet } = getSheetData();
  const headers = [
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
    "Signature Email"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  Logger.log("Sheet headers initialized");
}

// Test functions
function testQueue() {
  queue();
}

function testStatusUpdates() {
  updateAllDocumentStatuses();
}

function testSignature() {
  const { sheet, values } = getSheetData();
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row && row.length >= 4 && row[CONFIG.COLUMNS.STATUS] === CONFIG.STATUS.INITIAL_STAMPED) {
      processDocumentForSignature(sheet, i + 1, row);
      return;
    }
  }
  Logger.log("No documents with INITIAL_STAMPED status found");
}

