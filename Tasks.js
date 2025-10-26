const TASK_LIST_ID = 'MTY5ODIyNDk4MTk5MTQ5MjcxMjk6MDow'; // Replace this with your Google Tasks list ID
const SHEET_NAME = 'Tasks';
const SCRIPT_PROP_KEY = 'seen_task_ids';

function syncTasksToSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const seenTaskIds = getSeenTaskIds();
  const tasks = Tasks.Tasks.list(TASK_LIST_ID).items || [];
  const newTaskIds = [];

  tasks.forEach(task => {
    if (!seenTaskIds.includes(task.id)) {
      sheet.appendRow([task.title, task.notes || '', task.status, task.due || '', task.updated]);
      sendNewTaskEmail(task); // Send email when new task is found
      newTaskIds.push(task.id);
    }
  });

  if (newTaskIds.length > 0) {
    saveSeenTaskIds(seenTaskIds.concat(newTaskIds));
  }
}


function getSeenTaskIds() {
  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty(SCRIPT_PROP_KEY);
  return json ? JSON.parse(json) : [];
}

function saveSeenTaskIds(taskIds) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(SCRIPT_PROP_KEY, JSON.stringify(taskIds));
}

function sendNewTaskEmail(task) {
  var recipient = 'rbarrio1@alumni.nd.edu'; // Your email
  var subject = 'New Google Task: ' + task.title;
  var body = 'A new task has been synced to your sheet:\n\n' +
             'Title: ' + task.title + '\n' +
             'Notes: ' + (task.notes || 'None') + '\n' +
             'Status: ' + task.status + '\n' +
             'Due: ' + (task.due || 'None') + '\n' +
             'Updated: ' + task.updated;
  MailApp.sendEmail(recipient, subject, body);
}
