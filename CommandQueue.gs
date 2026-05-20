/**
 * AI Personal Assistant command queue bridge.
 *
 * Purpose:
 * - Let Ricky use ChatGPT as the UI.
 * - ChatGPT writes rows to COMMAND_QUEUE in the AI PERSONAL ASSISTANT Sheet.
 * - Apps Script processes complete_task commands.
 * - The original Google Task is marked complete by tasklist_id + task_id.
 * - Google Tasks keeps ownership of recurrence/repeating-task behavior.
 * - The Sheet re-syncs from Google Tasks and the Task Command Center Doc refreshes.
 *
 * Required advanced service:
 * - Tasks API advanced service enabled as `Tasks` in appsscript.json.
 */

const AIPA_CONFIG = Object.freeze({
  spreadsheetId: '1FVXQ2Xc7JBHGNmwwVKJT8XG1CTpyoWg6RzzWYmVXPrE',
  commandQueueSheet: 'COMMAND_QUEUE',
  tasksRawSheet: 'TASKS_RAW',
  openTasksSheet: 'OPEN_TASKS_FOR_CHATGPT',
  tasksSummarySheet: 'TASKS_SUMMARY',
  taskCommandCenterDocId: '1Fy0Md1a_N-CyTkBddSWRzaGGDh6Z6VY2ltV4fC9AUm0',
  maxOpenTasksForDoc: 12,
  maxTaskListsToSync: 100,
  maxTasksPerListPage: 100,
});

const AIPA_COMMAND_HEADERS = [
  'command_id',
  'created_at',
  'source',
  'user_text',
  'action',
  'matched_tasklist_id',
  'matched_task_id',
  'matched_title',
  'confidence',
  'status',
  'processed_at',
  'result',
  'error',
  'completed_at',
  'doc_refresh_status',
  'sync_status',
  'created_by',
  'notes',
];

const AIPA_TASKS_RAW_HEADERS = [
  'tasklist',
  'tasklist_id',
  'task_id',
  'title',
  'notes',
  'status',
  'due',
  'completed',
  'updated',
  'deleted',
  'hidden',
  'parent',
  'position',
  'task_link',
  'webViewLink',
  'etag',
  'last_synced_at',
];

const AIPA_OPEN_TASKS_HEADERS = [
  'tasklist',
  'tasklist_id',
  'task_id',
  'title',
  'notes',
  'due',
  'updated',
  'task_link',
  'webViewLink',
  'days_until_due',
];

/**
 * One-time setup. Run manually after deployment/authorization.
 * Creates/repairs the command queue and installs a time trigger.
 */
function setupAiPersonalAssistantCommandBridge() {
  ensureAipaCommandQueueSheet_();
  syncGoogleTasksToAiPersonalAssistantSheet();
  refreshTaskCommandCenterDocFromSheet();
  installAiPersonalAssistantCommandBridgeTrigger();
}

/**
 * Installs a 5-minute trigger to process command rows.
 */
function installAiPersonalAssistantCommandBridgeTrigger() {
  const handler = 'processAiPersonalAssistantCommandQueue';
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction && trigger.getHandlerFunction() === handler)
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyMinutes(5)
    .create();
}

/**
 * Main worker. Processes pending commands written by ChatGPT into COMMAND_QUEUE.
 */
function processAiPersonalAssistantCommandQueue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  const now = new Date();
  let syncStatus = 'not_run';
  let docStatus = 'not_run';

  try {
    const ss = SpreadsheetApp.openById(AIPA_CONFIG.spreadsheetId);
    const queue = ensureAipaCommandQueueSheet_(ss);
    const values = queue.getDataRange().getValues();
    if (values.length < 2) return;

    const headerMap = buildHeaderMap_(values[0]);
    let processedAny = false;

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const status = String(row[headerMap.status] || '').toLowerCase().trim();
      const action = String(row[headerMap.action] || '').toLowerCase().trim();

      if (status && status !== 'pending') continue;
      if (!action) continue;

      try {
        let result;
        if (action === 'complete_task') {
          result = completeTaskCommand_(ss, row, headerMap, now);
        } else {
          throw new Error('Unsupported command action: ' + action);
        }

        writeQueueResult_(queue, r + 1, headerMap, {
          status: 'processed',
          processed_at: now,
          result: result.message,
          error: '',
          completed_at: result.completedAt || now,
        });
        processedAny = true;
      } catch (err) {
        writeQueueResult_(queue, r + 1, headerMap, {
          status: 'error',
          processed_at: now,
          result: '',
          error: err && err.message ? err.message : String(err),
        });
      }
    }

    if (processedAny) {
      syncGoogleTasksToAiPersonalAssistantSheet();
      syncStatus = 'synced';
      refreshTaskCommandCenterDocFromSheet();
      docStatus = 'refreshed';
      stampLatestProcessedRows_(queue, headerMap, syncStatus, docStatus);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Convenience function ChatGPT can target by writing a command row into the Sheet.
 * This is also useful for testing manually inside Apps Script.
 */
function enqueueAiPersonalAssistantCompleteTask(userText, matchedTasklistId, matchedTaskId, matchedTitle, confidence) {
  const ss = SpreadsheetApp.openById(AIPA_CONFIG.spreadsheetId);
  const queue = ensureAipaCommandQueueSheet_(ss);
  const now = new Date();
  queue.appendRow([
    Utilities.getUuid(),
    now,
    'apps_script_manual',
    userText || '',
    'complete_task',
    matchedTasklistId || '',
    matchedTaskId || '',
    matchedTitle || '',
    confidence || '',
    'pending',
    '',
    '',
    '',
    '',
    '',
    '',
    Session.getActiveUser().getEmail() || 'unknown',
    '',
  ]);
}

function completeTaskCommand_(ss, row, headerMap, now) {
  let taskListId = String(row[headerMap.matched_tasklist_id] || '').trim();
  let taskId = String(row[headerMap.matched_task_id] || '').trim();
  let title = String(row[headerMap.matched_title] || '').trim();
  const userText = String(row[headerMap.user_text] || '').trim();

  if (!taskListId || !taskId) {
    const match = matchOpenTaskFromCommand_(ss, userText || title);
    if (!match) {
      throw new Error('Could not confidently match command to an open Google Task. Add matched_tasklist_id and matched_task_id, or make the command more specific.');
    }
    taskListId = match.taskListId;
    taskId = match.taskId;
    title = match.title;
  }

  const body = {
    status: 'completed',
    completed: now.toISOString(),
  };

  const updated = Tasks.Tasks.patch(body, taskListId, taskId);
  return {
    completedAt: now,
    message: 'Completed Google Task: ' + (title || updated.title || taskId),
  };
}

/**
 * Matches a loose command like "I finished Floss" to OPEN_TASKS_FOR_CHATGPT.
 * It only auto-matches when there is a single clear title match.
 */
function matchOpenTaskFromCommand_(ss, userText) {
  if (!userText) return null;

  const sheet = ss.getSheetByName(AIPA_CONFIG.openTasksSheet);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(h => String(h || '').trim());
  const col = name => headers.indexOf(name);
  const titleCol = col('title');
  const listIdCol = col('tasklist_id');
  const taskIdCol = col('task_id');
  const linkCol = col('task_link');
  const tasklistCol = col('tasklist');

  const query = normalizeTaskText_(userText);
  const matches = [];

  for (let i = 1; i < rows.length; i++) {
    const title = String(rows[i][titleCol] || '').trim();
    if (!title) continue;

    const normalizedTitle = normalizeTaskText_(title);
    const score = taskMatchScore_(query, normalizedTitle);
    if (score < 0.45) continue;

    let taskListId = listIdCol >= 0 ? String(rows[i][listIdCol] || '').trim() : '';
    let taskId = taskIdCol >= 0 ? String(rows[i][taskIdCol] || '').trim() : '';

    if ((!taskListId || !taskId) && linkCol >= 0) {
      const parsed = parseGoogleTaskApiLink_(String(rows[i][linkCol] || ''));
      taskListId = taskListId || parsed.taskListId;
      taskId = taskId || parsed.taskId;
    }

    if (taskListId && taskId) {
      matches.push({
        score,
        title,
        taskListId,
        taskId,
        tasklist: tasklistCol >= 0 ? String(rows[i][tasklistCol] || '') : '',
      });
    }
  }

  matches.sort((a, b) => b.score - a.score);
  if (!matches.length) return null;
  if (matches.length > 1 && Math.abs(matches[0].score - matches[1].score) < 0.12) return null;
  return matches[0];
}

/**
 * Full sync from Google Tasks into the assistant spreadsheet.
 * Recurrence remains owned by Google Tasks; this only mirrors current task state.
 */
function syncGoogleTasksToAiPersonalAssistantSheet() {
  const ss = SpreadsheetApp.openById(AIPA_CONFIG.spreadsheetId);
  const rawSheet = ensureSheetWithHeaders_(ss, AIPA_CONFIG.tasksRawSheet, AIPA_TASKS_RAW_HEADERS);
  const openSheet = ensureSheetWithHeaders_(ss, AIPA_CONFIG.openTasksSheet, AIPA_OPEN_TASKS_HEADERS);
  const summarySheet = ss.getSheetByName(AIPA_CONFIG.tasksSummarySheet) || ss.insertSheet(AIPA_CONFIG.tasksSummarySheet);

  const now = new Date();
  const taskLists = listAllGoogleTaskLists_();
  const rawRows = [];
  const openRows = [];

  taskLists.forEach(taskList => {
    const tasks = listAllGoogleTasksForList_(taskList.id);
    tasks.forEach(task => {
      const taskLink = 'https://www.googleapis.com/tasks/v1/lists/' + encodeURIComponent(taskList.id) + '/tasks/' + encodeURIComponent(task.id);
      rawRows.push([
        taskList.title || '',
        taskList.id || '',
        task.id || '',
        task.title || '',
        task.notes || '',
        task.status || '',
        task.due || '',
        task.completed || '',
        task.updated || '',
        Boolean(task.deleted),
        Boolean(task.hidden),
        task.parent || '',
        task.position || '',
        taskLink,
        task.webViewLink || '',
        task.etag || '',
        now,
      ]);

      if (isOpenTaskForChatGPT_(task)) {
        openRows.push([
          taskList.title || '',
          taskList.id || '',
          task.id || '',
          task.title || '',
          task.notes || '',
          task.due || '',
          task.updated || '',
          taskLink,
          task.webViewLink || '',
          daysUntilDue_(task.due, now),
        ]);
      }
    });
  });

  openRows.sort((a, b) => {
    const aDays = typeof a[9] === 'number' ? a[9] : 999999;
    const bDays = typeof b[9] === 'number' ? b[9] : 999999;
    if (aDays !== bDays) return aDays - bDays;
    return String(a[3]).localeCompare(String(b[3]));
  });

  replaceSheetData_(rawSheet, AIPA_TASKS_RAW_HEADERS, rawRows);
  replaceSheetData_(openSheet, AIPA_OPEN_TASKS_HEADERS, openRows);
  writeTasksSummary_(summarySheet, openRows, now);
}

function refreshTaskCommandCenterDocFromSheet() {
  const ss = SpreadsheetApp.openById(AIPA_CONFIG.spreadsheetId);
  const openSheet = ss.getSheetByName(AIPA_CONFIG.openTasksSheet);
  if (!openSheet) throw new Error('Missing sheet: ' + AIPA_CONFIG.openTasksSheet);

  const rows = openSheet.getDataRange().getValues();
  const headers = rows[0].map(h => String(h || '').trim());
  const col = name => headers.indexOf(name);
  const titleCol = col('title');
  const dueCol = col('due');
  const daysCol = col('days_until_due');
  const tasklistCol = col('tasklist');

  const data = rows.slice(1).filter(row => String(row[titleCol] || '').trim());
  const dueOrOverdue = data.filter(row => {
    const days = Number(row[daysCol]);
    return !isNaN(days) && days <= 0;
  });
  const nextRows = (dueOrOverdue.length ? dueOrOverdue : data).slice(0, AIPA_CONFIG.maxOpenTasksForDoc);

  const openCount = data.length;
  const overdueCount = data.filter(row => Number(row[daysCol]) < 0).length;
  const dueTodayCount = data.filter(row => Number(row[daysCol]) === 0).length;

  const lines = [];
  lines.push('CURRENT SUMMARY — AI PERSONAL ASSISTANT');
  lines.push('Updated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd h:mm a z'));
  lines.push('Purpose: This Doc is the short current summary for ChatGPT check-ins. The Sheets are the source of truth and detailed audit trail.');
  lines.push('');
  lines.push('SOURCE RULE');
  lines.push('- Use the AI PERSONAL ASSISTANT Sheet for the live task queue.');
  lines.push('- Use COMMAND_QUEUE for ChatGPT-to-Apps-Script task completion commands.');
  lines.push('- Use this Doc only as the current briefing, not as the raw task database.');
  lines.push('');
  lines.push('CURRENT STATUS');
  lines.push('- Live task sync shows ' + openCount + ' open tasks.');
  lines.push('- Overdue: ' + overdueCount + '. Due today: ' + dueTodayCount + '.');
  lines.push('- Main rule: surface due/overdue tasks before picking a next action.');
  lines.push('');
  lines.push('DUE / OVERDUE QUEUE TO SURFACE FIRST');
  if (!nextRows.length) {
    lines.push('No open tasks found in OPEN_TASKS_FOR_CHATGPT.');
  } else {
    nextRows.forEach((row, idx) => {
      const due = row[dueCol] ? String(row[dueCol]).slice(0, 10) : 'no due date';
      const tasklist = tasklistCol >= 0 ? String(row[tasklistCol] || '') : '';
      lines.push((idx + 1) + '. ' + row[titleCol] + ' — ' + due + (tasklist ? ' — ' + tasklist : ''));
    });
  }
  lines.push('');
  lines.push('NEXT TASK SUGGESTIONS BY MOTIVATION LEVEL');
  lines.push('Low motivation — under 5 minutes: Floss, GA / Journal, or write the next physical action for the top overdue item.');
  lines.push('Medium motivation — about 10 minutes: Do one small admin cleanup or clarify one overdue item.');
  lines.push('High motivation — about 30 minutes: Work the highest-value overdue tax/admin/client item until there is a clear completion or next blocker.');
  lines.push('');
  lines.push('OPERATING GUARDRAILS');
  lines.push('- No gambling.');
  lines.push('- No random fast food.');
  lines.push('- Use the diet plan instead of repeatedly editing the diet plan.');
  lines.push('- For repeating tasks, complete the original Google Task; do not recreate a duplicate task from the Sheet.');
  lines.push('');
  lines.push('MANUAL CHECK-IN COMMAND');
  lines.push('When Ricky says “Send check-in from Doc,” run an impromptu check-in immediately in the current conversation. Read the latest version of this Doc first, then cross-check the live AI PERSONAL ASSISTANT Sheet when available.');
  lines.push('');
  lines.push('TASK COMPLETION UPDATE RULE');
  lines.push('When Ricky says he finished or accomplished a task, write or process a COMMAND_QUEUE complete_task row, complete the original Google Task by tasklist_id + task_id, re-sync the Sheet, then refresh this Doc. If the exact task row is unclear, ask one targeted clarification or log it as an unlinked completion note instead of guessing.');

  const doc = DocumentApp.openById(AIPA_CONFIG.taskCommandCenterDocId);
  const body = doc.getBody();
  body.clear();
  body.setText(lines.join('\n'));
  doc.saveAndClose();
}

function ensureAipaCommandQueueSheet_(ss) {
  ss = ss || SpreadsheetApp.openById(AIPA_CONFIG.spreadsheetId);
  const sheet = ensureSheetWithHeaders_(ss, AIPA_CONFIG.commandQueueSheet, AIPA_COMMAND_HEADERS);
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureSheetWithHeaders_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const same = headers.every((h, i) => String(currentHeaders[i] || '') === h);
  if (!same) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function replaceSheetData_(sheet, headers, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.setFrozenRows(1);
}

function writeTasksSummary_(sheet, openRows, now) {
  const overdue = openRows.filter(row => Number(row[9]) < 0).length;
  const dueToday = openRows.filter(row => Number(row[9]) === 0).length;
  const dueThisWeek = openRows.filter(row => {
    const days = Number(row[9]);
    return !isNaN(days) && days >= 0 && days <= 7;
  }).length;

  const output = [
    ['metric', 'value'],
    ['last_sync', now],
    ['open_task_count', openRows.length],
    ['overdue_count', overdue],
    ['due_today_count', dueToday],
    ['due_this_week_count', dueThisWeek],
    ['', ''],
    ['top_open_tasks', ''],
    ['tasklist', 'title', 'due', 'days_until_due'],
  ];

  openRows.slice(0, 15).forEach(row => output.push([row[0], row[3], row[5], row[9]]));
  sheet.clearContents();
  sheet.getRange(1, 1, output.length, 4).setValues(output.map(row => {
    while (row.length < 4) row.push('');
    return row;
  }));
}

function listAllGoogleTaskLists_() {
  const result = [];
  let pageToken;
  do {
    const response = Tasks.Tasklists.list({
      maxResults: AIPA_CONFIG.maxTaskListsToSync,
      pageToken,
    });
    (response.items || []).forEach(item => result.push(item));
    pageToken = response.nextPageToken;
  } while (pageToken);
  return result;
}

function listAllGoogleTasksForList_(taskListId) {
  const result = [];
  let pageToken;
  do {
    const response = Tasks.Tasks.list(taskListId, {
      maxResults: AIPA_CONFIG.maxTasksPerListPage,
      pageToken,
      showCompleted: true,
      showDeleted: true,
      showHidden: true,
    });
    (response.items || []).forEach(item => result.push(item));
    pageToken = response.nextPageToken;
  } while (pageToken);
  return result;
}

function isOpenTaskForChatGPT_(task) {
  if (!task) return false;
  if (task.deleted) return false;
  if (task.hidden) return false;
  return task.status !== 'completed';
}

function daysUntilDue_(dueValue, now) {
  if (!dueValue) return '';
  const dueDateText = String(dueValue).slice(0, 10);
  const dueDate = new Date(dueDateText + 'T00:00:00');
  if (isNaN(dueDate.getTime())) return '';
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const target = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate());
  return Math.round((target.getTime() - today.getTime()) / 86400000);
}

function parseGoogleTaskApiLink_(link) {
  const result = { taskListId: '', taskId: '' };
  const match = String(link || '').match(/\/lists\/([^/]+)\/tasks\/([^/?#]+)/);
  if (!match) return result;
  result.taskListId = decodeURIComponent(match[1]);
  result.taskId = decodeURIComponent(match[2]);
  return result;
}

function normalizeTaskText_(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/\b(i|ive|i've|finished|finish|completed|complete|did|done|handled|sent|task|the|a|an|to|from|today)\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function taskMatchScore_(query, title) {
  if (!query || !title) return 0;
  if (query === title) return 1;
  if (title.indexOf(query) >= 0 || query.indexOf(title) >= 0) return 0.9;

  const qWords = new Set(query.split(' ').filter(Boolean));
  const tWords = new Set(title.split(' ').filter(Boolean));
  if (!qWords.size || !tWords.size) return 0;

  let overlap = 0;
  qWords.forEach(word => {
    if (tWords.has(word)) overlap++;
  });
  return overlap / Math.max(qWords.size, tWords.size);
}

function buildHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, idx) => {
    map[String(header || '').trim()] = idx;
  });
  AIPA_COMMAND_HEADERS.forEach(header => {
    if (!(header in map)) throw new Error('COMMAND_QUEUE missing header: ' + header);
  });
  return map;
}

function writeQueueResult_(sheet, rowNumber, headerMap, update) {
  Object.keys(update).forEach(key => {
    if (!(key in headerMap)) return;
    sheet.getRange(rowNumber, headerMap[key] + 1).setValue(update[key]);
  });
}

function stampLatestProcessedRows_(queue, headerMap, syncStatus, docStatus) {
  const values = queue.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][headerMap.status] || '').toLowerCase() === 'processed') {
      if (!values[r][headerMap.sync_status]) queue.getRange(r + 1, headerMap.sync_status + 1).setValue(syncStatus);
      if (!values[r][headerMap.doc_refresh_status]) queue.getRange(r + 1, headerMap.doc_refresh_status + 1).setValue(docStatus);
    }
  }
}
