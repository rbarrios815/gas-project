/***************************************
 * Google Voice → Sheet importer (voicemail + SMS)
 *
 * Behavior:
 * - Reads Gmail messages labeled "GV-To-Sheet"
 * - If it's a VOICEMAIL email:
 *     Extracts text between ".com>" and "Play message"
 *     (e.g., "Test 1 2 3") and writes that to Column A.
 * - If it's a TEXT MESSAGE email:
 *     Extracts the content between "<https://voice.google.com>"
 *     and "To respond to this text message" and writes only the
 *     main content line (e.g., "Test") to Column A.
 * - Leaves Column B blank.
 * - Inserts the email's received time into Column C
 *   on the same row, in MM/DD/YYYY HH:MM:SS format.
 * - Marks messages as processed using label "GV-Processed"
 *
 * No existing code/features in your other files are removed.
 ***************************************/
/***************************************
 * Google Voice → Sheet importer (voicemail + SMS)
 *
 * Behavior:
 * - Reads Gmail messages labeled "GV-To-Sheet"
 * - If it's a VOICEMAIL email:
 *     Extracts text between the first "https://voice.google.com"
 *     block and "play message", skipping URL lines.
 * - If it's a TEXT MESSAGE email:
 *     Extracts the first non-empty, non-URL line between
 *     "https://voice.google.com" and
 *     "To respond to this text message".
 * - Writes the extracted content into Column A.
 * - Leaves Column B blank.
 * - Inserts the email's received time into Column C
 *   on the same row, in MM/DD/YYYY HH:MM:SS format.
 * - Marks messages as processed using label "GV-Processed"
 ***************************************/

// ✅ Your spreadsheet ID (from the URL of your sheet)
const GV_SPREADSHEET_ID = '1rzejdmR0hatqESPp9MroCwT229QGM0oB2G9mELaL4Ps';

// Sheet tab name for gid=1564799966
const GV_SHEET_NAME = 'NOTES INBOX';

// Gmail label that marks Voice emails that should be imported
const GV_IMPORT_LABEL_NAME = 'GV-To-Sheet';

// Gmail label used to mark threads as already processed
const GV_PROCESSED_LABEL_NAME = 'GV-Processed';

// Only import messages from this phone number (via subject line)
const GV_ALLOWED_SUBJECTS = [
  'New voicemail from (281) 714-6370',
  'New text message from (281) 714-6370'
];

/**
 * Main function: call this from a time-based trigger.
 */
function importGoogleVoiceToSheet() {
  // ✅ Use explicit spreadsheet ID so it always writes to the correct file
  const ss = SpreadsheetApp.openById(GV_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(GV_SHEET_NAME);

  if (!sheet) {
    throw new Error(
      'Sheet with name "' + GV_SHEET_NAME + '" not found. ' +
      'Check that the tab name matches exactly.'
    );
  }

  // Cache existing notes in Column A to prevent duplicate entries.
  const existingNotes = buildExistingNotesSet_(sheet);

  // Get or create labels
  let importLabel = GmailApp.getUserLabelByName(GV_IMPORT_LABEL_NAME);
  if (!importLabel) {
    importLabel = GmailApp.createLabel(GV_IMPORT_LABEL_NAME);
  }

  let processedLabel = GmailApp.getUserLabelByName(GV_PROCESSED_LABEL_NAME);
  if (!processedLabel) {
    processedLabel = GmailApp.createLabel(GV_PROCESSED_LABEL_NAME);
  }

  // ✅ Search for threads with GV-To-Sheet but NOT yet GV-Processed
  const searchQuery =
    'label:"' + GV_IMPORT_LABEL_NAME + '" -label:"' + GV_PROCESSED_LABEL_NAME + '"';
  const threads = GmailApp.search(searchQuery).slice(0, 50); // limit per run

  if (!threads || threads.length === 0) {
    return; // Nothing to do
  }

  threads.forEach(function (thread) {
    const messages = thread.getMessages();
    messages.forEach(function (message) {
      const subject = message.getSubject() || '';
      if (!isAllowedVoiceSubject(subject)) {
        return; // Ignore messages from other numbers
      }

      const body = message.getPlainBody();
      if (!body) return;

      // Extract only the content we care about
      const extracted = extractVoiceContent(body);
      if (!extracted) return;

      const normalized = normalizeVoiceNote_(extracted);
      if (existingNotes.has(normalized)) {
        return; // Skip duplicates already in Column A
      }

      // Get the time the email was received
      const msgDate = message.getDate();
      const formattedDate = Utilities.formatDate(
        msgDate,
        Session.getScriptTimeZone(),
        'MM/dd/yyyy HH:mm:ss'
      );

      // FIRST EMPTY CELL IN COLUMN A
      const nextRow = sheet.getLastRow() + 1; // if 0 rows, becomes 1

      // Column A: extracted text (e.g., "Test 1 2 3" or "Test")
      sheet.getRange(nextRow, 1).setValue(extracted);
      existingNotes.add(normalized);

      // Column B: intentionally left blank

      // Column C: email received time in desired format
      sheet.getRange(nextRow, 3).setValue(formattedDate);
    });

    // Mark thread as processed: add processed label, remove import label
    thread.addLabel(processedLabel);
    thread.removeLabel(importLabel);
  });
}

function buildExistingNotesSet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const set = new Set();
  values.forEach(function (row) {
    const normalized = normalizeVoiceNote_(row[0]);
    if (normalized) set.add(normalized);
  });
  return set;
}

function normalizeVoiceNote_(note) {
  if (note === null || note === undefined) return '';
  return String(note).trim().replace(/\s+/g, ' ');
}

/**
 * Checks whether a subject matches the allowed Google Voice number.
 */
function isAllowedVoiceSubject(subject) {
  if (!subject) return false;
  var normalized = subject.trim();
  return GV_ALLOWED_SUBJECTS.some(function (allowed) {
    return normalized === allowed;
  });
}

/**
 * Extracts the relevant Google Voice content from the email body.
 *
 * Handles two patterns:
 * 1) Voicemail email:
 *    - Contains "play message"
 *    - Returns all non-empty, non-URL lines between
 *      the first "https://voice.google.com" and "play message".
 *
 * 2) Text message email:
 *    - Contains "To respond to this text message"
 *    - Returns the first non-empty, non-URL line between
 *      "https://voice.google.com" and
 *      "To respond to this text message".
 *
 * If no known pattern matches, returns null.
 */
function extractVoiceContent(body) {
  if (!body) return null;

  var trimmedBody = body.trim();
  var lower = trimmedBody.toLowerCase();

  var hasPlayMessage = lower.indexOf('play message') !== -1;
  var hasTextReply = lower.indexOf('to respond to this text message') !== -1;

  // Find first "https://voice.google.com"
  var startIdx = trimmedBody.indexOf('https://voice.google.com');
  if (startIdx === -1) {
    // Fallback: try ".com>" pattern you mentioned earlier
    var dotComIdx = trimmedBody.indexOf('.com>');
    if (dotComIdx !== -1) {
      startIdx = dotComIdx + '.com>'.length;
    } else {
      startIdx = 0;
    }
  }

  /**
   * Helper to turn a substring into cleaned lines:
   * - split by newline
   * - trim
   * - drop blank lines
   * - drop pure URLs
   */
  function cleanLines(textBlock) {
    return textBlock
      .split(/\r?\n/)
      .map(function (line) { return line.trim(); })
      .filter(function (line) {
        if (!line) return false;
        var clean = line.replace(/^<|>$/g, '');
        if (/^https?:\/\//i.test(clean)) return false; // skip pure URLs like <https://voice.google.com>
        return true;
      });
  }

  /**
   * TEXT MESSAGE pattern:
   * Body contains "to respond to this text message".
   * We take the first non-empty, non-URL line
   * between the voice.google.com link and that phrase.
   */
  if (hasTextReply) {
    var endIdxText = lower.indexOf('to respond to this text message');
    if (endIdxText > startIdx) {
      var innerText = trimmedBody.substring(startIdx, endIdxText);
      var contentTextLines = cleanLines(innerText);

      if (contentTextLines.length > 0) {
        // For SMS, just capture the main line like "Test"
        return contentTextLines[0];
      }
    }
  }

  /**
   * VOICEMAIL pattern:
   * Body contains "play message" (any case).
   * We take all non-empty, non-URL lines between
   * the voice.google.com link and "play message",
   * and join them into a single string.
   */
  if (hasPlayMessage) {
    var endIdxVm = lower.indexOf('play message', startIdx);
    if (endIdxVm === -1) {
      endIdxVm = lower.indexOf('play message');
    }

    if (endIdxVm > startIdx) {
      var innerVm = trimmedBody.substring(startIdx, endIdxVm);
      var contentVmLines = cleanLines(innerVm);

      if (contentVmLines.length > 0) {
        // Join into a single line for Column A
        return contentVmLines.join(' ');
      }
    }
  }

  // If we can't confidently parse, skip this email
  return null;
}

/**
 * Manual test helper.
 * Run this from the Script Editor to confirm behavior.
 */
function testImportGoogleVoiceToSheet() {
  importGoogleVoiceToSheet();
}
