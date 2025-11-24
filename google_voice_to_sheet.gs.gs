/***************************************
 * Google Voice → Sheet importer
 * - Reads Gmail messages labeled "GV-To-Sheet"
 * - Appends each message body into the first empty
 *   cell in Column A of a specific sheet
 * - Marks messages as processed using label "GV-Processed"
 *
 * No existing code/features in your project are removed.
 ***************************************/

// 1) CHANGE THIS to the exact name of the sheet tab
// that corresponds to gid=1564799966.
// Example: "NOTES INBOX", "Voice Inbox", etc.
const GV_SHEET_NAME = 'NOTES INBOX';

// 2) Gmail label that marks Voice emails that should be imported
const GV_IMPORT_LABEL_NAME = 'GV-To-Sheet';

// 3) Gmail label used to mark threads as already processed
const GV_PROCESSED_LABEL_NAME = 'GV-Processed';

/**
 * Main function: call this from a time-based trigger.
 * It will:
 *  - Find all Gmail threads with label GV-To-Sheet
 *  - For each message, append the plain-body text to the
 *    first empty row in Column A of the GV_SHEET_NAME sheet
 *  - Then move the thread to a "processed" label
 */
function importGoogleVoiceToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GV_SHEET_NAME);

  if (!sheet) {
    throw new Error(
      'Sheet with name "' + GV_SHEET_NAME + '" not found. ' +
      'Open the script and change GV_SHEET_NAME to your actual tab name.'
    );
  }

  // Get or create labels
  let importLabel = GmailApp.getUserLabelByName(GV_IMPORT_LABEL_NAME);
  if (!importLabel) {
    importLabel = GmailApp.createLabel(GV_IMPORT_LABEL_NAME);
  }

  let processedLabel = GmailApp.getUserLabelByName(GV_PROCESSED_LABEL_NAME);
  if (!processedLabel) {
    processedLabel = GmailApp.createLabel(GV_PROCESSED_LABEL_NAME);
  }

  // Grab up to 50 threads per run that have the import label
  // (You can increase this, but 50 is usually safe for quotas)
  const threads = importLabel.getThreads(0, 50);
  if (!threads || threads.length === 0) {
    // Nothing to do
    return;
  }

  threads.forEach(function (thread) {
    // Skip if already processed
    const hasProcessed = thread.getLabels().some(function (label) {
      return label.getName() === GV_PROCESSED_LABEL_NAME;
    });
    if (hasProcessed) {
      return;
    }

    const messages = thread.getMessages();
    messages.forEach(function (message) {
      const body = message.getPlainBody(); // plain text body
      if (!body) return;

      const cleaned = body.trim();
      if (!cleaned) return;

      // FIRST EMPTY CELL IN COLUMN A
      const nextRow = sheet.getLastRow() + 1; // if 0 rows, this becomes 1
      sheet.getRange(nextRow, 1).setValue(cleaned);
    });

    // Mark thread as processed: add processed label, remove import label
    thread.addLabel(processedLabel);
    thread.removeLabel(importLabel);
  });
}

/**
 * OPTIONAL: helper to test with a single run.
 * Manually run testImportGoogleVoiceToSheet() from the editor
 * to confirm it appends data as expected.
 */
function testImportGoogleVoiceToSheet() {
  importGoogleVoiceToSheet();
}
