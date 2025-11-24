function doGet(e) {
  var userEmail = Session.getActiveUser().getEmail(); // Ensure it gets the active user
  Logger.log("Detected User Email: " + userEmail); // Debugging - logs detected email

  var allowedUsers = ['jbgreatfamily1@gmail.com', 'rbarrios815@gmail.com', 'domlozano7@gmail.com', 'rbarrio1nd@gmail.com', 'rbarrio1@alumni.nd.edu','barriosgreatfamily1@gmail.com'];

  if (allowedUsers.includes(userEmail)) {
    ensureJbChipDailyTrigger();
    return HtmlService.createHtmlOutputFromFile('Index');
  } else {
    return HtmlService.createHtmlOutput("Sorry, you do not have access to this app.<br>Your detected email: " + userEmail);
  }
}




function getClientNamesAndCategories() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();
  var clientsByCategory = {
        '⭐⭐⭐⭐⭐': [],
    '⭐⭐⭐⭐': [],
    '⭐⭐⭐': [],
    'FUTURE FOLLOW UP SCHEDULED': [],
    '⭐⭐': [],
    '⭐': [],
    'MAINTENANCE': [],
    'PENDING ACCOUNT GROWTH': [],
    'NONE': [],
        'DONE': [] // Added "DONE" category here


  };

  // Process each row and organize clients by category
  data.forEach(function(row) {
    var clientName = row[0];
    var category = row[5];
    if (clientName && category && clientsByCategory.hasOwnProperty(category)) {
      clientName = clientName.replace(/\d+$/, '').trim(); // Remove numbers from the end of the client name
      // Check if clientName with category is already added to avoid duplicates when client has multiple rows
      var fullNameWithCategory = clientName + ' - ' + category;
      if (!clientsByCategory[category].includes(fullNameWithCategory)) { // Change logic to use fullNameWithCategory for uniqueness
        clientsByCategory[category].push(fullNameWithCategory);
      }
    }
  });

  // Sort clients within each category if needed
  for (var category in clientsByCategory) {
    clientsByCategory[category].sort(); // Sorts the full names with categories
  }

  // Select three clients from 'OPPORTUNITY' and 'MAINTENANCE' categories
  // var selectedOpportunityClients = clientsByCategory['OPPORTUNITY'].slice(0, 3);
  // var selectedMaintenanceClients = clientsByCategory['MAINTENANCE'].slice(0, 3);

  // Combine sorted clients into one array
  var combinedClients = []
          .concat(clientsByCategory['⭐⭐⭐⭐⭐'])
      .concat(clientsByCategory['⭐⭐⭐⭐'])
    .concat(clientsByCategory['⭐⭐⭐'])
    .concat(clientsByCategory['⭐⭐'])
    .concat(clientsByCategory['FUTURE FOLLOW UP SCHEDULED'])
    .concat(clientsByCategory['⭐'])
    .concat(clientsByCategory['MAINTENANCE'])
    .concat(clientsByCategory['PENDING ACCOUNT GROWTH'])
    .concat(clientsByCategory['NONE'])
        .concat(clientsByCategory['DONE']); // Include "DONE" category here



  // Since combinedClients now already include the category, return this array directly
  return combinedClients; // Now this array will have the client names with categories
}




////////////////////////////////////////////////////////
// REPLACE YOUR OLD getClientDetails(...) WITH THIS:
////////////////////////////////////////////////////////
function getClientDetails(clientName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();
  var clientNotes = [];
  var clientLabels = [];
  var columnB = "";
  var columnD = "";
  var columnLValue = "";

  // NEW: chip fields (P/Q)
  var chipInitials = ""; // Column P (index 15)
  var chipDateRaw = "";  // Column Q (index 16)

  data.forEach(function(row) {
    var thisClientName = row[0];
    if (thisClientName) {
      thisClientName = thisClientName.replace(/\d+$/, '').trim();
      if (thisClientName.toLowerCase() === clientName.toLowerCase()) {
        var thisCategory = row[5] ? row[5].toString().trim() : "N/A";
        var thisNoteDate = row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "MM/dd/yy") : "N/A";
        var thisFollowUpDate = row[3] ? Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "MM/dd/yy") : "N/A";
        var thisNote = row[2] ? row[2].toString().trim() : "N/A";
        var thisFollowUp = row[4] ? row[4].toString().trim() : "N/A";

        columnB = row[1];
        columnD = row[3];
        columnLValue = row[11];

        // Capture P/Q (may be blank)
        chipInitials = row[15] ? row[15].toString().trim() : "";
        chipDateRaw  = row[16] ? row[16] : ""; // may be Date or string

        clientNotes.push({
          clientName: thisClientName,
          noteDate: thisNoteDate,
          note: thisNote,
          followUpDate: thisFollowUpDate,
          followUp: thisFollowUp,
          category: thisCategory
        });

        // Build labels once for the first matching row, regardless of whether G is empty
if (clientLabels.length === 0) {

  // Column G (green, removable)
  if (row[6] && String(row[6]).trim() !== '') {
    var columnGLabels = String(row[6]).trim().split(' • ').map(function(s){ return s.trim(); }).filter(function(s){ return s.length; });
    columnGLabels.forEach(function(label) {
      clientLabels.push({ label: label, isBlue: false });
    });
  }

  // Columns H–K (blue, read-only)
  for (var i = 7; i <= 10; i++) {
    if (row[i] && String(row[i]).trim() !== '') {
      clientLabels.push({ label: String(row[i]).trim(), isBlue: true });
    }
  }
}

      }
    }
  });

  clientNotes.sort(function(a, b) {
    var date1 = new Date(a.noteDate);
    var date2 = new Date(b.noteDate);
    return date1 - date2;
  });

  // If chipDateRaw is a Date, format to MM/dd/yy; if string, try to parse; else blank.
  var chipDate = "";
  if (chipDateRaw) {
    var asDate = (chipDateRaw instanceof Date) ? chipDateRaw : new Date(chipDateRaw);
    if (!isNaN(asDate.getTime())) {
      chipDate = Utilities.formatDate(asDate, Session.getScriptTimeZone(), "MM/dd/yy");
    } else {
      chipDate = chipDateRaw.toString();
    }
  }

  return {
    notes:  clientNotes.length > 0 ? clientNotes : null,
    labels: clientLabels,
    columnB: columnB,
    columnD: columnD,
    columnL: columnLValue,
    chipInitials: chipInitials, // Column P
    chipDate: chipDate          // Column Q -> "MM/dd/yy" or ""
  };
}



// Extract and sort by date functions remain the same
function extractDate(note) {
  var datePattern = /^\d{2}\/\d{2}\/\d{2}/; // Matches MM/DD/YY at the beginning of the string
  var match = note.match(datePattern);
  return match ? new Date(match[0]) : new Date(0); // Default to a very early date if no match
}



// function getClientNotes(clientName) {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
//   var data = sheet.getDataRange().getValues();
//   var clientNotes = [];

//   // Loop through the data and push notes for clients with matching names
//   data.forEach(function(row, index) {
//     var thisClientName = row[0];
//     if (thisClientName) {
//       thisClientName = thisClientName.replace(/\d+$/, '').trim(); // Remove numbers from the end
//       if (thisClientName.toLowerCase() === clientName.toLowerCase()) {
//         var thisCategory = row[5] ? row[5].toString().trim() : "N/A";
//       var dateOptions = { timeZone: "GMT", formatType: "short" };
//       var thisNoteDate = row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "MM/dd/yy") : "N/A";
//     var thisFollowUpDate = row[3] ? Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "MM/dd/yy") : "N/A";
//         var thisNote = row[2] ? row[2].toString().trim() : "N/A";
//         var thisFollowUp = row[4] ? row[4].toString().trim() : "N/A";
        
//         clientNotes.push({
//           clientName: thisClientName,
//           noteDate: thisNoteDate,
//           note: thisNote,
//           followUpDate: thisFollowUpDate,
//           followUp: thisFollowUp,
//           category: thisCategory
//         });
//       }
//     }
//   });

//   // If clientNotes is empty, return null to trigger the 'no notes' message
//   return clientNotes.length > 0 ? clientNotes : null;


// function extractDate(note) {
//     var datePattern = /^\d{2}\/\d{2}\/\d{2}/; // Matches MM/DD/YY at the beginning of the string
//     var match = note.match(datePattern);
//     return match ? new Date(match[0]) : new Date(0); // Default to a very early date if no match
//   }

//   // Custom comparator for sorting by dates
//   function compareNotes(note1, note2) {
//     var date1 = extractDate(note1);
//     var date2 = extractDate(note2);
//     return date1 - date2;
//   }

//   // Sort the clientNotes array
//   clientNotes.sort(compareNotes);

//   // clientNotes is now sorted and can be used elsewhere in the program
// }

function updateClientNote(clientName, newNote) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var currentClientName = data[i][0].replace(/\s\d+$/, '').trim();
    if (currentClientName === clientName) {
      var currentFollowUp = data[i][4];
      var updatedFollowUp = currentFollowUp ? currentFollowUp + "\n" + newNote : newNote;
      sheet.getRange(i + 1, 5).setValue(updatedFollowUp);
      break;
    }
  }
}
function updatePastWork(clientName, pastWorkContent) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();
  var updated = false;
  var lastRowToUpdate = -1; // Initialize to an invalid row number
  var today = new Date();
  var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd/yy"); // Format the date as MM/dd/yy

  // Iterate over each row to find the last occurrence of the client's name
  for (var i = 1; i < data.length; i++) {
    var thisClientName = data[i][0];
    if (thisClientName) {
      thisClientName = thisClientName.replace(/\d+$/, '').trim(); // Remove numbers from the end
      var sheetClientName = clientName.trim().toLowerCase();
      var thisClientNameLowerCase = thisClientName.trim().toLowerCase();

      if (thisClientNameLowerCase === sheetClientName) {
        lastRowToUpdate = i; // Update the row number to the current row
      }
    }
  }

  // Check if a valid row was found and update the past work
  if (lastRowToUpdate != -1) {
    var pastWorkCell = sheet.getRange(lastRowToUpdate + 1, 3); // Column C for past work
    var existingPastWork = pastWorkCell.getValue();
    // Append new past work content with today's date
    var updatedPastWork = existingPastWork ? existingPastWork + "\n" + dateString + ": " + pastWorkContent : dateString + ": " + pastWorkContent;
    pastWorkCell.setValue(updatedPastWork);
    pastWorkCell.setBackground('yellow'); // Set the cell background to yellow
    updated = true;
  }

  return updated; // Return the status of the update operation
}

function updateFollowUp(clientName, newFollowUpContent, futureWorkDate, userSelection) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    var data = sheet.getDataRange().getValues();
    var updated = false;
    var lastRowToUpdate = -1; // Initialize to an invalid row number
    var dateString = "";

    // Process the date for both "JB" and "Ricky"
    if ((userSelection === 'JB' || userSelection === 'Ricky') && futureWorkDate) {
        var date = new Date(futureWorkDate);
        date.setDate(date.getDate() + 1);
        dateString = Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yy") + ": ";
        newFollowUpContent = "(" + userSelection + ")- " + newFollowUpContent; // Prepend "(JB)- " or "(Ricky)- " to the follow-up content
    }

    // Iterate over each row to find the last occurrence of the client's name
    for (var i = 1; i < data.length; i++) {
        var thisClientName = data[i][0];
        if (thisClientName) {
            thisClientName = thisClientName.replace(/\d+$/, '').trim(); // Remove numbers from the end
            var sheetClientName = clientName.trim().toLowerCase();
            var thisClientNameLowerCase = thisClientName.trim().toLowerCase();

            if (thisClientNameLowerCase === sheetClientName) {
                lastRowToUpdate = i; // Update the row number to the current row
            }
        }
    }

    // Check if a valid row was found and update the follow-up note
    if (lastRowToUpdate != -1) {
        var followUpCell = sheet.getRange(lastRowToUpdate + 1, 5);
        var existingFollowUp = followUpCell.getValue();
        // Append new follow-up content with the updated date or user tag
        var updatedFollowUp = existingFollowUp ? existingFollowUp + "\n" + dateString + newFollowUpContent : dateString + newFollowUpContent;
        followUpCell.setValue(updatedFollowUp);
        followUpCell.setBackground('yellow'); // Set the cell background to yellow
        updated = true;
    }

    return updated; // Return the status of the update operation
}







function showNotesForSelectedClient() {
  var clientName = document.getElementById('clientDropdown').value.trim();
  if (!clientName) return;

  // ✅ Add this line:
  onClientSelectedForBar(clientName, clientName);

  google.script.run
    .withSuccessHandler(displayClientNotes)
    .getClientNotes(clientName);
}


function displayClientNotes(notes) {
  if (notes) {
    // Code to display the notes on the page
  } else {
    // Code to display 'No notes available' or similar message
  }

}
function updateCategory(clientName, newCategory) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();
  var updated = false;

  for (var i = 1; i < data.length; i++) {
    var thisClientName = data[i][0];
    if (thisClientName) {
      thisClientName = thisClientName.replace(/\d+$/, '').trim(); // Remove numbers from the end
      if (thisClientName.toLowerCase() === clientName.toLowerCase()) {
        sheet.getRange(i + 1, 6).setValue(newCategory); // Update column F with the new category
        updated = true;
        // break; // Exit loop after updating to avoid unnecessary iterations
      }
    }
  }

  return updated;
}

// Add this in your script tag
function showCurrentCategory(clientCategory) {
  var categorySelect = document.getElementById('categorySelect');
  categorySelect.value = clientCategory; // Set the current category as the selected option
}
function getTimeSensitiveTasks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();
  var timeSensitiveTasks = [];
  var today = new Date();
  var fiveDaysBefore = new Date(today);
  fiveDaysBefore.setDate(today.getDate() - 10);
  var fiveDaysAfter = new Date(today);
  fiveDaysAfter.setDate(today.getDate() + 10);
  
  // Function to check if string contains a date in MM/dd/yy format
  function containsValidDate(str) {
    if (typeof str !== 'string') {
      return false;
    }

    var datePattern = /\d{2}\/\d{2}\/\d{2}/g; // Regular expression for MM/dd/yy format
    var matches = str.match(datePattern);
    if (matches) {
      for (var i = 0; i < matches.length; i++) {
        var parsedDate = Utilities.parseDate(matches[i], Session.getScriptTimeZone(), "MM/dd/yy");
        if (parsedDate && parsedDate >= fiveDaysBefore && parsedDate <= fiveDaysAfter) {
          return true;
        }
      }
    }
    return false;
  }

  // Check each row for time-sensitive dates
  data.forEach(function(row) {
    var clientName = row[0];
    var category = row[5]; // Assuming the category is in column F
    var noteDate = row[1] ? new Date(row[1]) : null;
    var followUpDate = row[3] ? new Date(row[3]) : null;
    var note = row[2];
    var followUpNote = row[4];

    // Include only OPPORTUNITY and MAINTENANCE categories
    if (category !== 'OPPORTUNITY' && category !== 'MAINTENANCE') {
      return; // Skip to the next iteration
    }

    // Format date and check if it's within time-sensitive range, and include category
    if (clientName && noteDate && noteDate >= fiveDaysBefore && noteDate <= fiveDaysAfter) {
      var formattedNoteDate = Utilities.formatDate(noteDate, Session.getScriptTimeZone(), "MM/dd/yy");
      timeSensitiveTasks.push(clientName + '- ' + category + ' - ' + formattedNoteDate + ' : ' + note);
    }
    if (clientName && followUpDate && followUpDate >= fiveDaysBefore && followUpDate <= fiveDaysAfter) {
      var formattedFollowUpDate = Utilities.formatDate(followUpDate, Session.getScriptTimeZone(), "MM/dd/yy");
      timeSensitiveTasks.push(clientName + '- ' + category + ' - ' + formattedFollowUpDate + ' : ' + followUpNote);
    }
    // Check if followUpNote contains a valid date in the range, and include category
    if (clientName && followUpNote && containsValidDate(followUpNote)) {
      timeSensitiveTasks.push(clientName + '- ' + category + '-' + followUpNote);
    }
  });

  return timeSensitiveTasks;
}



function addNewClient(clientName, note, category) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var lastRow = sheet.getLastRow();
  var nextRow = lastRow + 1; // Get next empty row
  
  // Set the new client's name, note, and category
  sheet.getRange(nextRow, 1).setValue(clientName); // Update column A with client name
  sheet.getRange(nextRow, 5).setValue(note).setBackground('yellow'); // Update column E with note
  sheet.getRange(nextRow, 6).setValue(category); // Update column F with category

  // Optionally, return some value to confirm the operation was successful
  return "Client added successfully";
}
function getClientNames() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    var data = sheet.getRange('A2:A' + sheet.getLastRow()).getValues(); // Adjust range as needed
    var clientNames = data.map(function(row) { return row[0]; });
    return clientNames.filter(function(name) { return name !== ''; }); // Filter out empty names
}

function getClientsByCategory(category) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    var data = sheet.getDataRange().getValues();
    var clients = [];

    data.forEach(function(row) {
        if (row[5] === category) { // Assuming category is in column F
            clients.push({
                name: row[0], // Assuming name is in column A
                date: row[1], // Assuming date is in column B
                note: row[2], // Assuming note is in column C
                followUp: row[4], // Assuming follow-up is in column E
                category: row[5]  // Assuming category is in column F
            });
        }
    });

    return clients;
}

// ///////ADD TASK TO CALENDAR
// function addEventToCalendar(clientName, calendarDate, calendarNote) {
//       if (!calendarDate) {
//         return 'Error: No date selected for the task.';
//     }
//     var calendar = CalendarApp.getCalendarById('rbarrios815@gmail.com');
    
//     var eventDate = new Date(calendarDate);
//     eventDate.setDate(eventDate.getDate() + 1);

//     var event;
//     try {
//         event = calendar.createAllDayEvent(clientName, eventDate, {description: calendarNote});
//         if (event) {
//                       var formattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "MM/dd/yy"); // Changed format to MM/dd/yy

//             return 'Task for ' + clientName + ' on ' + formattedDate + ' added to calendar. CONFIRM STRICT/FLEXIBLE DROPDOWN AND DAY OF WEEK DROPDOWN';
//         }
//     } catch (e) {
//         Logger.log('Error creating event: ' + e.toString());
//         return 'Failed to add task. Error: ' + e.toString();
//     }

//     return 'Failed to add task.';
// }

// function addEventToRickyCalendar(clientName, calendarDate, calendarNote) {
//       if (!calendarDate) {
//         return 'Error: No date selected for the task.';
//     }
//     var calendar = CalendarApp.getCalendarById('9f797f1cac9f9ef80526c493150d461cae595535cafa9204f16b93e2eff8d446@group.calendar.google.com');
    
//     var eventDate = new Date(calendarDate);
//     eventDate.setDate(eventDate.getDate() + 1);

//     var event;
//     try {
//         event = calendar.createAllDayEvent(clientName, eventDate, {description: calendarNote});
//         if (event) {
//                       var formattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "MM/dd/yy"); // Changed format to MM/dd/yy

//             return 'Task for ' + clientName + ' on ' + formattedDate + ' added to calendar. DOUBLE CHECK DAY-OF-WEEK DROPDOWN & STRICT/FLEXIBLE DROPDOWN';
//         }
//     } catch (e) {
//         Logger.log('Error creating event: ' + e.toString());
//         return 'Failed to add task. Error: ' + e.toString();
//     }

//     return 'Failed to add task.';
// }
function addAppointmentToCalendar(clientName, appointmentDate, appointmentTime) {
    if (!appointmentDate || !appointmentTime) {
        return 'Error: Appointment date and/or time not selected.';
    }
    var calendar = CalendarApp.getCalendarById('rbarrios815@gmail.com');
    var timeZone = calendar.getTimeZone(); // Get the calendar's time zone

    // Parse the date and time in the calendar's time zone
    var startDate = Utilities.parseDate(appointmentDate + ' ' + appointmentTime, timeZone, 'yyyy-MM-dd HH:mm');
    var endDate = new Date(startDate.getTime() + 60 * 60 * 1000); // Add 1 hour for the appointment duration

    try {
    var event = calendar.createEvent(clientName + ' - APPOINTMENT', startDate, endDate);
    if (event) {
        // Format date as MM/DD and time as HH:MM AM/PM
        var formattedDate = Utilities.formatDate(startDate, timeZone, "MM/dd/YY hh:mm a");

        // Update follow-up note with exact appointment details in the desired format
        var followUpNote = "APPOINTMENT SCHEDULED FOR " + formattedDate;
        updatePastWork(clientName, followUpNote); // Assuming updateClientNote is a function to update the follow-up note

        // Return confirmation message
        return 'Appointment for ' + clientName + ' at ' + formattedDate + ' added to calendar and follow-up note updated.';
    }
} catch (e) {
    Logger.log('Error creating appointment: ' + e.toString());
    return 'Failed to add appointment. Error: ' + e.toString();
}

    return 'Failed to add appointment.';
}

function getTopClients() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();

  var clientMap = {};
  var categorySortOrder = {
    'TOP PRIORITY': 1
    // (optional) add other categories if you want priority within sort
  };

  var tz = Session.getScriptTimeZone();

  data.forEach(function(row, index) {
    if (index === 0) return; // header

    var rawName = row[0];
    if (!rawName) return;

    // Normalize name (strip trailing digits like "Acme 2")
    var clientName = rawName.toString().replace(/\d+$/, '').trim();

    var columnB       = row[1];   // STRICT/FLEX or a date, as in your sheet
    var pastWorkNote  = row[2];   // Column C
    var columnD       = row[3];   // Day-of-week (Mon/Tues/etc) or a date
    var followUpNote  = row[4];   // Column E
    var category      = row[5];   // Column F
    var columnLRaw    = row[11];  // Column L (In Progress)
    var chipInitials  = row[15] ? row[15].toString().trim() : ""; // Column P
    var chipDateRaw   = row[16];                                 // Column Q

    // 🔒 Normalize Column L to a safe string (front-end filters on this!)
    var colL = (columnLRaw == null) ? '' : columnLRaw.toString();

    // Format chip date as MM/dd/yy when present
    var chipDate = "";
    if (chipDateRaw) {
      var d = (chipDateRaw instanceof Date) ? chipDateRaw : new Date(chipDateRaw);
      if (!isNaN(d.getTime())) {
        chipDate = Utilities.formatDate(d, tz, "MM/dd/yy");
      } else {
        chipDate = chipDateRaw.toString();
      }
    }

    var key = clientName + ':' + (category || ''); // combine name+category

    if (!clientMap[key]) {
      clientMap[key] = {
        name: clientName,
        category: category,
        columnB: columnB,
        columnD: columnD,
        followUps: [],
        pastWorks: [],
        sortOrder: categorySortOrder[category] || Number.MAX_SAFE_INTEGER,

        // ✅ Expose Column L under BOTH names for the UI
        columnL: colL,
        columnLContent: colL,

        chipInitials: chipInitials,
        chipDate: chipDate
      };
      if (followUpNote) clientMap[key].followUps.push(followUpNote);
      if (pastWorkNote) clientMap[key].pastWorks.push(pastWorkNote);
    } else {
      if (followUpNote) clientMap[key].followUps.push(followUpNote);
      if (pastWorkNote) clientMap[key].pastWorks.push(pastWorkNote);

      // Keep first non-empty Column L, else update if we find a non-empty later
      if (!clientMap[key].columnL && colL) {
        clientMap[key].columnL = colL;
        clientMap[key].columnLContent = colL;
      }

      // Prefer the latest non-empty chip fields (override with current if present)
      if (chipInitials) clientMap[key].chipInitials = chipInitials;
      if (chipDate)     clientMap[key].chipDate     = chipDate;

      // Keep columnB/columnD if empty previously
      if (!clientMap[key].columnB && columnB) clientMap[key].columnB = columnB;
      if (!clientMap[key].columnD && columnD) clientMap[key].columnD = columnD;
    }
  });

  var topClients = Object.values(clientMap).sort(function(a, b) {
    return a.sortOrder - b.sortOrder || a.name.localeCompare(b.name);
  });

  return topClients;
}









function postponeTeamTask(rowIndex, newDate) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    var dateCell = sheet.getRange(rowIndex, 7);
    var newDateObj = new Date(newDate);
    newDateObj.setDate(newDateObj.getDate() + 1); // Add 1 to the date
    dateCell.setValue(newDateObj);
    return Utilities.formatDate(newDateObj, Session.getScriptTimeZone(), "MM/dd/yy");
}

function deleteTeamTask(rowIndex) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    sheet.deleteRow(rowIndex + 1); // Adding 1 because array index starts at 0, but Sheets rows start at 1
}


function getTaskDetail(rowIndex) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    var taskDetail = sheet.getRange(rowIndex + 1, 8).getValue(); // Assuming column H is the 8th column
    return taskDetail;
}

function getRickyTasks() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    var data = sheet.getDataRange().getValues();
    var rickyTasks = [];

    Logger.log("Total rows in data: " + data.length);

    data.forEach(function(row, index) {
        // Logging each row for debugging
        Logger.log("Row " + index + ": " + row.join(", "));

        var clientName = row[0];
        var noteDate = row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "MM/dd/yy") : "N/A";
        var note = row[2] || "N/A";
        var followUpDate = row[3] ? Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "MM/dd/yy") : "N/A";
        var followUp = row[4] || "N/A";
        var category = row[5] || "N/A";

        // Check if any cell in the row contains 'RICKY' and if the category is not 'FUTURE FOLLOW UP SCHEDULED'
        if (row.join(" ").toUpperCase().includes("RICKY") && category !== 'FUTURE FOLLOW UP SCHEDULED') {
            // Formatting the task for display
            var taskFormatted = clientName + " (" + category + "):\n" +
                                noteDate + "- " + note + "\n" +
                                followUpDate + "- " + followUp;

            rickyTasks.push(taskFormatted);
        }
    });

    Logger.log("Ricky Tasks found: " + rickyTasks.length);
    return rickyTasks;
}






// Google Apps Script

function getTeamTasks3() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    var data = sheet.getDataRange().getValues();
    var teamTasks3 = [];

    data.forEach(function(row, index) {
        if (row[0].startsWith('TEAM TASK')) {
            var taskParts = row[0].split(" - ", 2); // Split the task string
            var displayTask = taskParts.length > 1 ? taskParts[1] : ""; // Keep everything after the first dash
            teamTasks3.push({ task: displayTask, rowIndex: index + 1 });
        }
    });

    return teamTasks3;
}


function updateTeamTask3Content(rowIndex, newTaskContent) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    var updatedContent = "TEAM TASK - " + newTaskContent; // Prepend "TEAM TASK - " to the new content
    sheet.getRange(rowIndex, 1).setValue(updatedContent); // Update the cell in the spreadsheet
}

// Server-side Google Apps Script


function getDomTasks() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    var data = sheet.getDataRange().getDisplayValues();
    var clientGroups = {};
    var domTasks = [];
    var today = new Date();

    // Group entries by client name if in 'FUTURE FOLLOW UP SCHEDULED' category
    data.forEach(function(row, index) {
        if (index === 0 || row[5] !== 'FUTURE FOLLOW UP SCHEDULED') return; // Skip header row and non-relevant categories

        var clientName = row[0].replace(/\d+$/, '').trim(); // Remove numbers from end of client name
        if (!(clientName in clientGroups)) {
            clientGroups[clientName] = [];
        }
        clientGroups[clientName].push(row);
    });

    // Evaluate each group
    for (var client in clientGroups) {
        var hasFutureDate = clientGroups[client].some(function(row) {
            return checkForFutureDate(row[4], today);
        });

        if (!hasFutureDate) {
            var clientDetails = clientGroups[client].map(function(row) {
                return row.slice(0, 6).join(" - ");
            });
            domTasks = domTasks.concat(clientDetails);
        }
    }

    return domTasks;
}

function checkForFutureDate(text, today) {
    var datePattern = /\b\d{1,2}\/\d{1,2}\/\d{2,4}\b/g; // Regex to find date in MM/DD/YY or MM/DD/YYYY format
    var hasFutureDate = false;
    var matches = text.match(datePattern);
    
    if (matches) {
        matches.forEach(function(dateStr) {
            var dateParts = dateStr.split("/");
            var year = dateParts[2].length === 2 ? "20" + dateParts[2] : dateParts[2];
            var date = new Date(year, dateParts[0] - 1, dateParts[1]);

            console.log("Checking date:", dateStr, "Parsed Date:", date, "Today:", today);
            if (date >= today) {
                hasFutureDate = true;
                console.log("Future date found:", dateStr);
            }
        });
    }

    return hasFutureDate;
}








function addNewClientWithPastWork(clientName, note, category, pastWork) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    
    // Add new client with past work
    var lastRow = sheet.getLastRow();
    var nextRow = lastRow + 1; // Get next empty row

    sheet.getRange(nextRow, 1).setValue(clientName); // Column A: Client Name
    sheet.getRange(nextRow, 5).setValue(note); // Column E: Note
    sheet.getRange(nextRow, 6).setValue(category); // Column F: Category
    sheet.getRange(nextRow, 3).setValue(pastWork); // Column C: Past Work

    // Apply formula to the corresponding row in column L
    var formulaCell = sheet.getRange('P' + nextRow); // Target cell in column P for the new row
    var formula = '=IF(TRIM(F' + nextRow + ')="⭐⭐⭐⭐⭐", 0,' +
                  'IF(TRIM(F' + nextRow + ')="⭐⭐⭐⭐", 2,' +
                  'IF(TRIM(F' + nextRow + ')="⭐⭐⭐", 3,' +
                  'IF(TRIM(F' + nextRow + ')="⭐⭐", 1,' +
                  'IF(TRIM(F' + nextRow + ')="CLIENT OF THE DAY", 2,' +
                  'IF(TRIM(F' + nextRow + ')="⭐", 5,' +
                  'IF(TRIM(F' + nextRow + ')="PENDING ACCOUNT GROWTH", 6,' +
                  'IF(TRIM(F' + nextRow + ')="", 10,' +
                  'IF(TRIM(F' + nextRow + ')="MAINTENANCE", 7,' +
                  'IF(TRIM(F' + nextRow + ')="FUTURE FOLLOW UP SCHEDULED", 9, 8))))))))))';
    formulaCell.setFormula(formula);

    return "Client added and formula applied successfully.";
}



function getCalendarTasksFromPast14Days() {
  var today = new Date();
  var fourteenDaysAgo = new Date();
  fourteenDaysAgo.setDate(today.getDate() - 14);
  
  // Set end time of the last day to include full day
  today.setHours(23, 59, 59, 999);

  var calendarId = 'rbarrios815@gmail.com';
  var calendar = CalendarApp.getCalendarById(calendarId);

  // Retrieve all events in one call
  var events = calendar.getEvents(fourteenDaysAgo, today);
  var tasks = [];

  events.forEach(function(event) {
    if (event.isAllDayEvent()) {
      var title = event.getTitle();
      var clientName = title; // Assuming the title is the client's name
      var category = getCategoryForClient(clientName); // Get the category for the client

      tasks.push({
        clientName: clientName,
        date: Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        description: event.getDescription() || "No description",
        category: category || "Category Not Found",
        eventId: event.getId()
      });
    }
  });

  return tasks;
}







function getCategoryForClient(clientName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0'); // Replace with your actual sheet name
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
if (data[i][0].toLowerCase() === clientName.toLowerCase()) {
            return data[i][5]; // Returns the category from Column F
        }
    }
    return null; // Return null if the client name is not found
}



function getClientCategoriesForPieChart() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();
  var clientCategories = {};

  // Assuming the category is in column F and client names are in column A
  for (var i = 1; i < data.length; i++) {
    var clientName = data[i][0].trim();
    var category = data[i][5].trim();
    
    // Check if the client name is already accounted for
    if (!clientCategories[clientName]) {
      // Initialize if not present
      clientCategories[clientName] = category;
    }
  }

  // Count the categories
  var categoryCounts = {};
  for (var client in clientCategories) {
    var cat = clientCategories[client];
    categoryCounts[cat] = (categoryCounts[cat] || 0) + 1;
  }

  // Prepare data for pie chart
  var chartData = [['Category', 'Number of Clients']];
  for (var cat in categoryCounts) {
    chartData.push([cat, categoryCounts[cat]]);
  }

  return chartData;
}

function getWeekdayInteractionCounts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();

  var weekdayCounts = { 'Sunday': 0, 'Monday': 0, 'Tuesday': 0, 'Wednesday': 0, 'Thursday': 0, 'Friday': 0, 'Saturday': 0 };

  data.forEach(function(row, index) {
    if (index === 0) return; // Skip header row

    var noteDateColumnB = row[1]; // Assuming dates are in column B
    var noteDateColumnC = extractDateFromText(row[2]); // Assuming column C has text with date

    incrementWeekdayCount(weekdayCounts, noteDateColumnB);
    incrementWeekdayCount(weekdayCounts, noteDateColumnC);
  });

  return weekdayCounts;
}

function extractDateFromText(text) {
  var datePattern = /\d{2}\/\d{2}\/\d{2,4}/; // Adjust regex pattern as needed
  var match = text.match(datePattern);
  return match ? new Date(match[0]) : null;
}

function incrementWeekdayCount(weekdayCounts, date) {
  if (date) {
    var noteDate = new Date(date);
    if (!isNaN(noteDate.getTime())) { // Check if the date is valid
      var weekday = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][noteDate.getDay()];
      weekdayCounts[weekday]++;
    }
  }
}


function getWeekdayCountsForHeatmap() {
    var counts = getWeekdayInteractionCounts();
    return counts;
}

function getFutureWeekdayInteractionCounts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();

  var weekdayCounts = { 'Sunday': 0, 'Monday': 0, 'Tuesday': 0, 'Wednesday': 0, 'Thursday': 0, 'Friday': 0, 'Saturday': 0 };

  data.forEach(function(row, index) {
    if (index === 0) return; // Skip header row

    var futureDateColumnD = row[3]; // Assuming future dates are in column D
    var futureDateColumnE = extractDateFromText(row[4]); // Assuming column E has text with date

    incrementWeekdayCount(weekdayCounts, futureDateColumnD);
    incrementWeekdayCount(weekdayCounts, futureDateColumnE);
  });

  return weekdayCounts;
}
function getFutureWeekdayCountsForHeatmap() {
  var counts = getFutureWeekdayInteractionCounts();
  return counts;
}

function extractDateFromText(text) {
  if (text === null || text === undefined) return null;
  text = String(text); // Convert to string
  var datePattern = /\d{2}\/\d{2}\/\d{2,4}/;
  var match = text.match(datePattern);
  return match ? new Date(match[0]) : null;
}

// Function to get data for scatter plot
function getDataForScatterPlot() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    var data = sheet.getDataRange().getValues();
    var interactionCounts = {}; // Object to hold interaction counts per category

    // Loop through data and count interactions
    for (var i = 1; i < data.length; i++) {
        var category = data[i][5]; // Assuming category is in column F
        // Count dates in columns B, D, and within text in columns C, E
        var interactions = countDatesInRow(data[i]);
        if (!interactionCounts[category]) {
            interactionCounts[category] = 0;
        }
        interactionCounts[category] += interactions;
    }

    // Convert to array format for Google Charts
    var chartData = [['Category', 'Number of Interactions']];
    for (var category in interactionCounts) {
        chartData.push([category, interactionCounts[category]]);
    }
    return chartData;
}

// Helper function to count dates in a row (implement your logic to count dates)
function countDatesInRow(row) {
    var dateCount = 0;

    // Check if the row itself is undefined or null
    if (!row) {
        return dateCount; // Return 0 if the row is undefined or null
    }

    // Check columns B and D for dates
    [1, 3].forEach(function(index) {
        // Check if the cell is not undefined or null
        if (row[index] && new Date(row[index]).toString() !== 'Invalid Date') {
            dateCount++;
        }
    });

    // Regular expression to match dates in MM/DD/YYYY format
    var dateRegex = /\b\d{1,2}\/\d{1,2}\/\d{4}\b/g;

    // Check columns C and E for dates in text
    [2, 4].forEach(function(index) {
        // Check if the cell is not undefined or null
        if (row[index]) {
            var matches = row[index].toString().match(dateRegex);
            if (matches) {
                dateCount += matches.length;
            }
        }
    });

    return dateCount;
}






function generateTable() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('INCOME METERS / MONTH OVER MONTH');
    var range = sheet.getRange('A1:G33');
    var values = range.getValues();
    var backgrounds = range.getBackgrounds();

    return {values: values, backgrounds: backgrounds};
}


function getBirthdayClients() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    var data = sheet.getDataRange().getValues();
    var today = new Date();
    var currentMonth = today.getMonth();
    var birthdays = {};

    data.forEach(function(row) {
        var clientStatus = row[14]; // Column O for client status
        var birthdayStr = row[13]; // Column N for birthday string
        var clientName = row[0]; // Column A for client names
        if (clientStatus === 'YES' && birthdayStr) {
            var birthday = new Date(birthdayStr);
            var birthdayMonth = birthday.getMonth();
            if (isInRelevantMonth(birthdayMonth, currentMonth)) {
                var monthLabel = getMonthLabel(birthdayMonth);
                if (!birthdays[monthLabel]) {
                    birthdays[monthLabel] = [];
                }
                birthdays[monthLabel].push({name: clientName, birthday: birthday, birthdayStr: birthdayStr});
            }
        }
    });

    // Sort the birthdays within each month in chronological order and format the date
    for (var month in birthdays) {
        birthdays[month].sort(function(a, b) {
            return a.birthday - b.birthday;
        }).map(function(client) {
            client.birthday = Utilities.formatDate(client.birthday, Session.getScriptTimeZone(), "MM/dd");
            return client;
        });
    }

    return birthdays;
}


function isInRelevantMonth(birthdayMonth, currentMonth) {
    var relevantMonths = [];
    if (currentMonth === 0) { // January
        relevantMonths = [11, 0, 1]; // Dec, Jan, Feb
    } else if (currentMonth === 11) { // December
        relevantMonths = [10, 11, 0]; // Nov, Dec, Jan
    } else {
        relevantMonths = [currentMonth - 1, currentMonth, currentMonth + 1];
    }
    return relevantMonths.includes(birthdayMonth);
}

function getMonthLabel(monthIndex) {
    var monthNames = ["January", "February", "March", "April", "May", "June",
                      "July", "August", "September", "October", "November", "December"];
    return monthNames[monthIndex];
}




function getTasksForDay(dayOffset) {
  var targetDate = new Date();
  targetDate.setDate(targetDate.getDate() + dayOffset);
  var formattedDate = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Get clients in 'TOP PRIORITY' and 'APPOINTMENT SCHEDULED' categories
  var clients = getClientsByCategories(['⭐⭐⭐⭐⭐', '⭐⭐⭐⭐'], formattedDate);

  // Get calendar events based on the day
  var calendarEvents = getCalendarEventsForDay(formattedDate, dayOffset);

  return { calendarEvents: calendarEvents, clients: clients };
}

function getCalendarEventsForDay(formattedDate, dayOffset) {
  switch (dayOffset) {
    case -1: // Yesterday
      return getYesterdaysCalendarEvents(formattedDate);
    case 0: // Today
      return get10DaysCalendarEvents(formattedDate);
    case 1: // Tomorrow
      return getTomorrowsCalendarEvents(formattedDate);
    default:
      return []; // Default case if none of the above
  }
}



function getClientsByCategories(categories) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();
  var clients = [];

  data.forEach(function(row) {
    var clientName = row[0];
    var category = row[5];
    if (clientName && categories.includes(category)) {
      clientName = clientName.replace(/\d+$/, '').trim(); // Remove trailing numbers
clients.push(clientName + ' - ' + category);
    }
  });

  return [...new Set(clients)]; // Remove duplicates
}

function getYesterdaysCalendarEvents() {
  var calendarId = 'rbarrios815@gmail.com'; // Your calendar ID
  var calendar = CalendarApp.getCalendarById(calendarId);

  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1); // Set to yesterday
  var startOfDay = new Date(yesterday.getFullYear(), yesterday.getMonth(), yesterday.getDate());
  var endOfDay = new Date(startOfDay);
  endOfDay.setDate(endOfDay.getDate() + 1);
  endOfDay.setMilliseconds(-1); // Set to the last millisecond of yesterday

  var events = calendar.getEvents(startOfDay, endOfDay);
  var calendarEvents = [];

  events.forEach(function(event) {
    var startTime = event.getStartTime();
    var endTime = event.getEndTime();

    var formattedStartTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'HH:mm');
    var formattedEndTime = Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'HH:mm');

    if (event.isAllDayEvent()) {
      calendarEvents.push('TASK: ' + event.getTitle() + ' - ' + (event.getDescription() || 'No description'));
    } else {
      calendarEvents.push('APPOINTMENT: ' + event.getTitle() + ' - ' + formattedStartTime + ' to ' + formattedEndTime + ' - ' + (event.getDescription() || 'No description'));
    }
  });

  return calendarEvents;
}

function get10DaysCalendarEvents() {
    var calendarId = 'rbarrios815@gmail.com'; // Your calendar ID
    var calendar = CalendarApp.getCalendarById(calendarId);

    var today = new Date();
    var startOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    var endOf30Days = new Date(startOfDay);
    endOf30Days.setDate(endOf30Days.getDate() + 10); // Set to 30 days from today

    var events = calendar.getEvents(startOfDay, endOf30Days);
    var calendarEvents = [];

    events.forEach(function(event) {
        var startTime = event.getStartTime();
        var endTime = event.getEndTime();
        
        var formattedStartTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
        var formattedEndTime = Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

        if (event.isAllDayEvent()) {
            calendarEvents.push('TASK: ' + event.getTitle() + ' - ' + (event.getDescription() || 'No description'));
        } else {
            calendarEvents.push('APPOINTMENT: ' + event.getTitle() + ' - ' + formattedStartTime + ' to ' + formattedEndTime + ' - ' + (event.getDescription() || 'No description'));
        }
    });

    return calendarEvents;
}


function getTomorrowsCalendarEvents() {
  var calendarId = 'rbarrios815@gmail.com'; // Your calendar ID
  var calendar = CalendarApp.getCalendarById(calendarId);

  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1); // Set to tomorrow
  var startOfDay = new Date(tomorrow.getFullYear(), tomorrow.getMonth(), tomorrow.getDate());
  var endOfDay = new Date(startOfDay);
  endOfDay.setDate(endOfDay.getDate() + 1);
  endOfDay.setMilliseconds(-1); // Set to the last millisecond of tomorrow

  var events = calendar.getEvents(startOfDay, endOfDay);
  var calendarEvents = [];

  events.forEach(function(event) {
    var startTime = event.getStartTime();
    var endTime = event.getEndTime();

    var formattedStartTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'HH:mm');
    var formattedEndTime = Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'HH:mm');

    if (event.isAllDayEvent()) {
      calendarEvents.push('TASK: ' + event.getTitle() + ' - ' + (event.getDescription() || 'No description'));
    } else {
      calendarEvents.push('APPOINTMENT: ' + event.getTitle() + ' - ' + formattedStartTime + ' to ' + formattedEndTime + ' - ' + (event.getDescription() || 'No description'));
    }
  });

  return calendarEvents;
}








function getAllLabels() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  var labelsSet = new Set();

  data.forEach(function(row) {
    var labelsCell = row[6]; // Assuming labels are in column G
    if (labelsCell) {
      var labels = labelsCell.toString().split(' • ');
      labels.forEach(label => labelsSet.add(label.trim()));
    }
  });

  var labelsArray = Array.from(labelsSet);
  labelsArray.sort(); // Sort labels alphabetically
  return labelsArray;
}




function addLabelToClient(clientName, label) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();
  var clientFound = false;

  for (var i = 0; i < data.length; i++) {
    var thisClientName = data[i][0]; // Assuming client names are in column A
    if (thisClientName && thisClientName.toString().toLowerCase() === clientName.toLowerCase()) {
      clientFound = true;
      var existingLabels = data[i][6] ? data[i][6].toString() : ''; // Column G for labels
      var newLabels = existingLabels ? existingLabels + " • " + label : label;
      sheet.getRange(i + 1, 7).setValue(newLabels); // Update the cell in column G
      break;
    }
  }

  if (!clientFound) {
    throw new Error('Client not found');
  }
}

function removeLabelFromClient(clientName, labelToRemove) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    var data = sheet.getDataRange().getValues();

    for (var i = 0; i < data.length; i++) {
        if (data[i][0].trim().toLowerCase() === clientName.toLowerCase()) {
            var currentLabels = data[i][6] ? data[i][6].toString().split(' • ') : [];
            var updatedLabels = currentLabels.filter(label => label.trim() !== labelToRemove.trim());
            sheet.getRange(i + 1, 7).setValue(updatedLabels.join(' • '));
            break;
        }
    }
}



function getTodayNotes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();
  var todayNotes = [];
  var datePattern = createDatePattern(new Date());

  data.forEach(function(row) {
    var clientName = row[0];
    var pastNotes = row[2]; // Column C
    var followUps = row[4]; // Column E

    // Function to process each note and add it to todayNotes if it contains today's date
    function processNote(note, isPastNote) {
      if (note.match(datePattern)) {
        // Add a flag to indicate if the note is from past notes (column C)
        todayNotes.push({ note: clientName + ': ' + note, isPastNote: isPastNote });
      }
    }

    if (clientName) {
      if (pastNotes) {
        pastNotes.split('\n').forEach(note => processNote(note, true)); // Pass true for past notes
      }
      if (followUps) {
        followUps.split('\n').forEach(note => processNote(note, false)); // Pass false for follow-ups
      }
    }
  });

  return todayNotes;
}

function createDatePattern(date) {
    let day = ('0' + date.getDate()).slice(-2);
    let month = ('0' + (date.getMonth() + 1)).slice(-2);
    let year2d = date.getFullYear().toString().substr(-2);
    let year4d = date.getFullYear().toString();

    // Patterns for different date formats
    let patterns = [
        `${month}/${day}/${year2d}`, // MM/DD/YY
        `${month}/${day}/${year4d}`, // MM/DD/YYYY
        `${parseInt(month)}/${day}/${year2d}`, // M/DD/YY
        `${parseInt(month)}/${day}/${year4d}`, // M/DD/YYYY
        `${month}/${parseInt(day)}/${year2d}`, // MM/D/YY
        `${month}/${parseInt(day)}/${year4d}`, // MM/D/YYYY
        `${parseInt(month)}/${parseInt(day)}/${year2d}`, // M/D/YY
        `${parseInt(month)}/${parseInt(day)}/${year4d}`  // M/D/YYYY
    ];

    // Combine into a single regex pattern
    return new RegExp(patterns.join('|'), 'g');
}


function updateClientData(clientName, newData) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
    var data = sheet.getDataRange().getValues();

    for (var i = 0; i < data.length; i++) {
        if (data[i][0].trim() === clientName.trim()) { // Assuming column A has client names
            sheet.getRange(i + 1, 12).setValue(newData); // Column L is the 12th column
            break;
        }
    }
}

 /* Update Column L ("In Progress") for the given client.
 * Mirrors updateInProgress but exposes a concise name for the client UI.
 * Returns a small payload so the frontend can trust the saved value.
 */
function updateClientColumnL(clientName, newValue) {
  var name = String(clientName || '').trim();
  if (!name) {
    throw new Error('Missing clientName.');
  }

  var text = newValue == null ? '' : String(newValue);

  // Reuse the normalized lookup logic that already strips trailing digits.
  updateInProgress(name, text);

  // Keep recent-clients metadata in sync (ignore failures – best effort only).
  try {
    logRecentClient(name, 5);
  } catch (err) {
    console.warn('logRecentClient failed inside updateClientColumnL:', err);
  }

  return {
    ok: true,
    client: name,
    value: text
  };
}


// Example of function to handle quadrant selection
function selectQuadrant(clientType, timeFrame) {
    // Assuming you have a function to call your Google Apps Script with parameters
    google.script.run.withSuccessHandler(function(result) {
        // Handle the display of filtered and sorted clients
        showTopClients(result);
    }).getTopClientsFiltered(clientType, timeFrame);
}

function updateClientColumnB(clientName, columnBValue) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues(); // Get all data from the sheet

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === clientName) { // Assuming client names are in the first column
      var row = i + 1; // Sheet rows are 1-indexed, arrays are 0-indexed
      var range = sheet.getRange(row, 2); // Column B is the second column
      range.setValue(columnBValue); // Set the new value for Column B
      break; // Exit the loop once the client is found and updated
    }
  }
}

function updateClientColumnD(clientName, columnDValue) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues(); // Get all data from the sheet

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === clientName) { // Assuming client names are in the first column
      var row = i + 1; // Sheet rows are 1-indexed, arrays are 0-indexed
      var range = sheet.getRange(row, 4); // Column D is the fourth column
      range.setValue(columnDValue); // Set the new value for Column D
      break; // Exit the loop once the client is found and updated
    }
  }
}

function getTopClientsByDay(day) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();
  var filteredClients = [];

  data.forEach(function(row, index) {
    if (index === 0) return; // Skip header row
    var columnD = row[3]; // Assuming this is the day column
    if (columnD.includes(day)) {
      var clientName = row[0].replace(/\d+$/, '').trim();
      var latestFollowUp = getLatestFollowUp(row[4]); // Assuming follow ups are in column E
      filteredClients.push({name: clientName, latestFollowUp: latestFollowUp});
    }
  });
  return filteredClients;
}

function getLatestFollowUp(followUps) {
  var latestDate = null;
  var latestNote = '';
  followUps.split('\n').forEach(function(note) {
    var dates = note.match(/\d{1,2}\/\d{1,2}\/\d{2,4}/);
    if (dates) {
      var date = new Date(dates[0]);
      if (!latestDate || date > latestDate) {
        latestDate = date;
        latestNote = note;
      }
    }
  });
  return latestNote;
}

function updateNotes(clientName, type, newText, originalText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0'); // Replace with your sheet name
  var data = sheet.getDataRange().getValues();
  var changes = [];

  if (type === 'pastWork') {
    changes = deriveNoteChanges_(originalText, newText);
  }

  // Find the row for the client
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === clientName) { // Assuming client names are in the first column
      if (type === 'pastWork') {
        sheet.getRange(i + 1, 3).setValue(newText); // Replace 2 with the correct column number for past work
        if (changes.length) {
          var synced = syncDashboardChangesToInbox_(clientName, changes);
          if (!synced) {
            throw new Error('EDIT FAILED ON OTHER VIEW');
          }
        }
      } else if (type === 'followUp') {
        sheet.getRange(i + 1, 5).setValue(newText); // Replace 3 with the correct column number for follow-ups
      }
      return;
    }
  }
  throw new Error('Client not found');
}

function deriveNoteChanges_(oldText, newText) {
  var oldLines = parsePastWorkLines_(oldText);
  var newLines = parsePastWorkLines_(newText);
  var max = Math.max(oldLines.length, newLines.length);
  var changes = [];

  for (var i = 0; i < max; i++) {
    var oldNote = oldLines[i] ? oldLines[i].noteText : '';
    var newNote = newLines[i] ? newLines[i].noteText : '';
    if (!oldNote || !newNote) continue;
    if (oldNote !== newNote) {
      changes.push({ oldNote: oldNote, newNote: newNote });
    }
  }
  return changes;
}

function parsePastWorkLines_(text) {
  return String(text || '')
    .split('\n')
    .map(function(line) {
      return {
        raw: line,
        noteText: line.replace(/^\s*\d{1,2}\/\d{1,2}\/\d{2,4}\s*:\s*/, '').trim()
      };
    });
}

function syncDashboardChangesToInbox_(clientName, changes) {
  var sh = ensureNotesInbox_();
  var last = sh.getLastRow();
  if (last < 2) return changes.length === 0;

  var data = sh.getRange(2, 1, last - 1, 2).getValues();
  var target = normalizeClientName_(clientName);
  var pending = changes.map(function(c) {
    return { old: String(c.oldNote || '').trim(), updated: String(c.newNote || '').trim(), matched: false };
  });

  for (var i = 0; i < data.length; i++) {
    var assigned = normalizeClientName_(data[i][1]);
    if (!assigned || assigned !== target) continue;

    var currentNote = String(data[i][0] || '').trim();
    for (var j = 0; j < pending.length; j++) {
      if (pending[j].matched) continue;
      if (currentNote === pending[j].old) {
        sh.getRange(i + 2, 1).setValue(pending[j].updated);
        pending[j].matched = true;
        break;
      }
    }
  }

  return pending.every(function(p) { return p.matched; });
}


function sendEditSummaryEmail(clientName, type, oldText, newText) {
  var recipient = 'rbarrio1@alumni.nd.edu'; // Replace with your email
  var subject = 'Edit Summary for ' + clientName;
  var body = 'Edits made to ' + type + ' for ' + clientName + ':\n\n' +
             'Old Text:\n' + oldText + '\n\n' +
             'New Text:\n' + newText;
  MailApp.sendEmail(recipient, subject, body);
}


function getAllClientsData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();
  var clients = [];

  for (var i = 1; i < data.length; i++) { // Start from 1 to skip header row if there's a header
    var row = data[i];
    var clientName = row[0];
    if (clientName) {
      var clientData = {
        clientName: clientName,
        category: row[5], // Assuming column F is index 5
        rowText: row.join(' ').toLowerCase() // Concatenate all cell values in the row
      };
      clients.push(clientData);
    }
  }
  return clients;
}

/** Return sorted unique, non-empty categories from Column F (row 2 → last). */
function getUniqueCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('DASHBOARD 8.0');
  const last = sh.getLastRow();
  if (last < 2) return [];

  const vals = sh.getRange(2, 6, last - 1, 1).getValues()  // Col F
                 .map(r => String(r[0] || '').trim())
                 .filter(Boolean);

  const uniq = Array.from(new Set(vals));
  uniq.sort(function(a, b){ return a.localeCompare(b, 'en', { sensitivity:'base' }); });
  return uniq;
}

/**
 * Update a client's Category (Column F) by exact match on Column A (Client Name).
 * Returns { ok:true } if saved, otherwise throws.
 */
function updateClientCategory(clientName, newCategory) {
  if (!clientName) throw new Error('Missing clientName.');
  if (!newCategory) throw new Error('Missing category.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('DASHBOARD 8.0');
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues(); // A..F

  let foundRow = -1;
  for (let i = 0; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    if (name.toLowerCase() === String(clientName).trim().toLowerCase()) {
      foundRow = i + 2; // sheet row (offset from header)
      break;
    }
  }

  if (foundRow === -1) {
    throw new Error('Client not found: ' + clientName);
  }

  // Write Column F
  sh.getRange(foundRow, 6).setValue(newCategory);

  return { ok: true, row: foundRow, category: newCategory };
}


function getClientsWithBirthdays() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sheet.getDataRange().getValues();
  var clientMap = {};
  for (var i = 1; i < data.length; i++) { // Start from 1 to skip header row
    var row = data[i];
    var clientName = row[0];
    var birthdayStr = row[13]; // Column N is index 13
    if (clientName) {
      var key = clientName;
      if (!clientMap[key]) {
        clientMap[key] = {
          name: clientName,
          birthday: null,
          month: null,
          day: null,
          columnB: row[1],
          columnD: row[3],
          category: row[5],
          followUps: [],
          pastWorks: [],
          sortOrder: 0 // Adjust as needed
        };
      }
      if (birthdayStr) {
        var dateParts = birthdayStr.split('/');
        if (dateParts.length === 2) {
          var month = parseInt(dateParts[0], 10);
          var day = parseInt(dateParts[1], 10);
          clientMap[key].birthday = new Date(new Date().getFullYear(), month - 1, day);
          clientMap[key].month = month;
          clientMap[key].day = day;
        }
      }
      // Collect followUps and pastWorks
      if (row[4]) clientMap[key].followUps.push(row[4]);
      if (row[2]) clientMap[key].pastWorks.push(row[2]);
    }
  }
  // Convert the map to an array
  var clients = Object.values(clientMap);

  // Now, sort the clients
  var currentMonth = new Date().getMonth() + 1; // Months are 0-indexed, so add 1
  clients.sort(function(a, b) {
    if (a.birthday && b.birthday) {
      var monthA = a.month >= currentMonth ? a.month : a.month + 12;
      var monthB = b.month >= currentMonth ? b.month : b.month + 12;
      if (monthA === monthB) {
        return a.day - b.day;
      }
      return monthA - monthB;
    } else if (a.birthday && !b.birthday) {
      return -1;
    } else if (!a.birthday && b.birthday) {
      return 1;
    } else {
      return 0;
    }
  });

  return clients;
}


function getCurrentMonthRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var currentDate = new Date();
  var currentMonth = currentDate.getMonth(); // 0-based (0 = January)
  var currentYear = currentDate.getFullYear();

  // Starting from Q2 for November 2024
  var startMonth = 10; // November (0-based)
  var startYear = 2024;
  var rowOffset = 2; // Q2

  // Calculate the difference in months from the start month
  var monthDifference = (currentYear - startYear) * 12 + (currentMonth - startMonth);

  if (monthDifference < 0) {
    throw new Error("Current date is before the start date of the audit trail.");
  }

  var targetRow = rowOffset + monthDifference;

  return targetRow;
}

function getAuditTrailMessages() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var targetRow;
  try {
    targetRow = getCurrentMonthRow();
  } catch (e) {
    Logger.log(e.message);
    return [];
  }

  var cell = sheet.getRange(targetRow, 17); // Column Q is the 17th column
  var cellValue = cell.getValue();

  if (!cellValue) {
    return [];
  }

  var messages = cellValue.split('\n').filter(msg => msg.trim() !== '');

  return messages;
}

function addAuditTrailMessage(message) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var targetRow;
  try {
    targetRow = getCurrentMonthRow();
  } catch (e) {
    Logger.log(e.message);
    return false;
  }

  var cell = sheet.getRange(targetRow, 17); // Column Q is the 17th column
  var existingMessages = cell.getValue();
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "M/d/yy h:mm a");
  var newMessage = formattedDate + " " + message;

  if (existingMessages) {
    cell.setValue(existingMessages + "\n" + newMessage);
  } else {
    cell.setValue(newMessage);
  }

  return true;
}

function submitMessageToAuditTrail(message) {
  if (!message || message.trim() === "") {
    throw new Error("Message cannot be empty.");
  }

  var wasAdded = addAuditTrailMessage(message.trim());
  return wasAdded;
}

function getAllAuditTrailMessages() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var lastRow = sheet.getLastRow();
  var messages = [];

  for (var row = 2; row <= lastRow; row++) { // Starting from Q2
    var cellValue = sheet.getRange(row, 17).getValue();
    if (cellValue) {
      var monthMessages = cellValue.split('\n').filter(msg => msg.trim() !== '');
      messages = messages.concat(monthMessages);
    }
  }

  return messages;
}


////////////////////////////////////////////////////////////////////////////////////////////////
// FIXED sendEmailNotification(clientName, followUpContent) to only email the single most
// recent line from multi-line Past Work (column C) and Future Work (column E).
////////////////////////////////////////////////////////////////////////////////////////////////
function sendEmailNotification(clientName, followUpContent) {
  // 1) Grab the client's details from getClientDetails()
  var details = getClientDetails(clientName);
  if (!details || !details.notes) {
    MailApp.sendEmail({
      to: 'rbarrio1@alumni.nd.edu',
      subject: 'New Dashboard Update Submitted',
      body: clientName + "\n\n(No client details found.)"
    });
    return;
  }

  // 2) From details, get columnB (STRICT/FLEXIBLE) and columnD (Day of Week)
  var strictOrFlexible = details.columnB ? details.columnB.toString().trim() : "";
  var dayOfWeek        = details.columnD ? details.columnD.toString().trim() : "";

  // 3) We'll track the single best Past Work line, and single best Future Work line
  var mostRecentPastLine    = null; // Will store the entire text of the best line
  var mostRecentPastDateObj = null; // Will store a Date for comparison

  var mostRecentFutureLine    = null;
  var mostRecentFutureDateObj = null;

  //////////////////////////////////////////
  // STEP A: Loop through all noteObj from details.notes
  //         But now we SPLIT noteObj.note and noteObj.followUp by "\n"
  //////////////////////////////////////////
  details.notes.forEach(function(noteObj) {
    
    // --- Past Work: noteObj.note might have multiple lines ---
    if (noteObj.note && noteObj.note.trim() !== "N/A") {
      var pastLines = noteObj.note.split("\n");
      pastLines.forEach(function(singleLine) {
        var lineDate = extractDateFromLine(singleLine);
        if (lineDate && (!mostRecentPastDateObj || lineDate > mostRecentPastDateObj)) {
          mostRecentPastDateObj = lineDate;
          mostRecentPastLine    = singleLine; // the entire text of that line
        }
      });
    }

    // --- Future Work: noteObj.followUp might have multiple lines ---
    if (noteObj.followUp && noteObj.followUp.trim() !== "N/A") {
      var futureLines = noteObj.followUp.split("\n");
      futureLines.forEach(function(singleLine) {
        var lineDate = extractDateFromLine(singleLine);
        if (lineDate && (!mostRecentFutureDateObj || lineDate > mostRecentFutureDateObj)) {
          mostRecentFutureDateObj = lineDate;
          mostRecentFutureLine    = singleLine;
        }
      });
    }
  });

  //////////////////////////////////////////
  // STEP B: If no valid line was found, fallback to "N/A"
  //////////////////////////////////////////
  if (!mostRecentPastLine) {
    mostRecentPastLine = "N/A";
  }
  if (!mostRecentFutureLine) {
    mostRecentFutureLine = "N/A";
  }

  //////////////////////////////////////////
  // STEP C: Format the lines to show the date + text
  //         We'll parse out the date from the line if it exists
  //////////////////////////////////////////
  var finalPastDateStr    = "N/A";
  var finalPastWorkUpdate = mostRecentPastLine; 
  var dateInPastLine      = extractDateFromLine(mostRecentPastLine);
  if (dateInPastLine) {
    // Remove that date from the line if you prefer, or keep it:
    // For clarity, let's keep it visible: "MM/dd/yy: rest of text"
    // But we can parse it out to do "date: rest"
    var parted = splitOutDate(mostRecentPastLine);
    finalPastDateStr    = parted.dateStr || "N/A";
    finalPastWorkUpdate = parted.rest.trim() || parted.dateStr; // if rest is empty, just show the date
  }

  var finalFutureDateStr    = "N/A";
  var finalFutureWorkUpdate = mostRecentFutureLine;
  var dateInFutureLine      = extractDateFromLine(mostRecentFutureLine);
  if (dateInFutureLine) {
    var partedFuture = splitOutDate(mostRecentFutureLine);
    finalFutureDateStr    = partedFuture.dateStr || "N/A";
    finalFutureWorkUpdate = partedFuture.rest.trim() || partedFuture.dateStr;
  }

  //////////////////////////////////////////
  // STEP D: Build the final email body
  //////////////////////////////////////////
  var emailBody =
    clientName + "\n\n" +
    strictOrFlexible + " : " + dayOfWeek + "\n\n" +
    finalPastDateStr + ": " + finalPastWorkUpdate + "\n\n" +
    finalFutureDateStr + ": " + finalFutureWorkUpdate;

  //////////////////////////////////////////
  // STEP E: Send the email
  //////////////////////////////////////////
  MailApp.sendEmail({
    to: "rbarrio1@alumni.nd.edu",
    subject: "New Dashboard Update Submitted",
    body: emailBody
  });
}


/**
 * Extracts a date object from a line of text (MM/DD/YY or MM/DD/YYYY).
 * Returns null if none found.
 */
function extractDateFromLine(line) {
  if (!line) return null;
  // e.g. /(\d{1,2})\/(\d{1,2})\/(\d{2,4})/
  var match = line.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (match) {
    var month = parseInt(match[1], 10) - 1;
    var day   = parseInt(match[2], 10);
    var year  = parseInt(match[3], 10);
    if (year < 100) {
      year = 2000 + year; 
    }
    var dt = new Date(year, month, day);
    if (!isNaN(dt.getTime())) {
      return dt;
    }
  }
  return null;
}



/**
 * Splits out the leading date from a line (if present).
 * Returns an object { dateStr: "MM/dd/yy...", rest: "the remainder" }.
 */
function splitOutDate(fullLine) {
  if (!fullLine) {
    return { dateStr: null, rest: "" };
  }
  var match = fullLine.match(/^(\d{1,2}\/\d{1,2}\/\d{2,4})(.*)/);
  if (match) {
    return { 
      dateStr: match[1].trim(), 
      rest: match[2] // the remainder after the date
    };
  }
  // if no leading date
  return { dateStr: null, rest: fullLine };
}



/**
 * Helper function to parse a string like "MM/dd/yy" or "MM/dd/yyyy" into a Date object.
 * Returns null if invalid.
 */
function tryParseDate(dateStr) {
  if (!dateStr) return null;
  // If dateStr is already recognized by JS as valid, let's do a simple parse:
  var parsed = new Date(dateStr);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }
  // If that fails, optionally handle custom patterns, but usually the above works
  return null;
}



function getStrictPastTopClients() {
  var clients = getTopClients(); // Retrieves an array of client objects
  var todayDay = new Date().toLocaleString('en-US', { weekday: 'short' }); // e.g., 'Mon', 'Tue', etc.

  var strictPastClients = clients.filter(function(client) {
    return client.columnB === 'STRICT' && client.columnD === todayDay;
  }).map(function(client) {
    return client.name;
  });

  return strictPastClients;
}
// Code.gs — REPLACE entire getUniqueCategories with this
function getUniqueCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DASHBOARD 8.0');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // no data rows

  // Column F (6th), rows 2..last
  const values = sheet
    .getRange(2, 6, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(v => (v == null ? '' : String(v).trim()))
    .filter(v => v !== ''); // drop blanks (including "" from formulas)

  if (values.length === 0) return [];

  // Case-insensitive dedupe while preserving first-seen casing
  const seen = new Map(); // key = lowercased, value = original
  for (const v of values) {
    const k = v.toLowerCase();
    if (!seen.has(k)) seen.set(k, v);
  }

  // Sort a→z, case-insensitive
  return Array.from(seen.values())
    .sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));
}




////////////////////////////////////////////////////////////////////////////////////////////////
// PASTE THIS SNIPPET AFTER THE LAST LINE OF YOUR Code.gs (e.g., after line 866)
////////////////////////////////////////////////////////////////////////////////////////////////

/******************************************************************************************************
 PASTE THIS ENTIRE FUNCTION IN YOUR Code.gs. IT WILL EMAIL THE “IMMEDIATE” (STRICT-PAST) CLIENTS EXACTLY
 AS YOUR HOMEPAGE DOES, SHOWING THEIR PAST NOTE, NEXT NOTE, AND COLUMN L (“IN PROGRESS”).
 ******************************************************************************************************/
function sendDailyTopClientsEmail() {
  // 1) Get the same allClients array from your existing getTopClients().
  //    Each client object looks like:
  //       { name, category, columnB, columnD, followUps: [...], pastWorks: [...], columnLContent, ... }
  var allClients = getTopClients();

  // 2) We replicate the exact “strictPast” logic:
  //    - "columnB === 'STRICT'"
  //    - “latest follow-up date” <= 7 days from now
  //    (In your code, you do 7 days by "fiftyHoursFromNow" or "new Date() + 7 days".)
  var now = new Date();
  var sevenDaysFromNow = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000); // 7 days
  // We must figure out each client’s “latestDate” from followUps so it matches your original code:
  allClients.forEach(function(client) {
    client.latestDate = findLargestDateAmong(client.followUps);
  });

  // Filter to “strictPast”
  var filteredClients = allClients.filter(function(client) {
    return (
      client.columnB === "STRICT" &&
      client.latestDate &&
      client.latestDate <= sevenDaysFromNow
    );
  });

  // 3) Next, we do the same sorting your showTopClients() does:
  //    a) Sort by latestDate ascending
  filteredClients.sort(function(a, b) {
    var dateA = a.latestDate ? a.latestDate.getTime() : 0;
    var dateB = b.latestDate ? b.latestDate.getTime() : 0;
    return dateA - dateB;
  });

  //    b) Sort by day-of-week in order starting from “today”
  var daysOfWeek = ["Sun","Mon","Tues","Wed","Thurs","Fri","Sat"];
  var todayIndex = new Date().getDay();
  var sortedDays = daysOfWeek.slice(todayIndex).concat(daysOfWeek.slice(0, todayIndex));
  filteredClients.sort(function(a, b) {
    var indexA = sortedDays.indexOf(a.columnD);
    var indexB = sortedDays.indexOf(b.columnD);
    if (indexA < 0) indexA = Infinity; // If no day-of-week, place last
    if (indexB < 0) indexB = Infinity;
    return indexA - indexB;
  });

  // 4) Separate those whose columnD = today from the “rest of the week”
  var todayDay = daysOfWeek[todayIndex];
  var clientsMatchingToday = [];
  var otherClients = [];
  filteredClients.forEach(function(client) {
    if (client.columnD === todayDay) {
      clientsMatchingToday.push(client);
    } else {
      otherClients.push(client);
    }
  });

  // 5) Build the email body
  var emailBody = "IMMEDIATE (strictPast) Clients:\n\n";

  // 5a) “TODAY” clients first
  if (clientsMatchingToday.length > 0) {
    emailBody += "TODAY: " + todayDay + "\n\n";
    clientsMatchingToday.forEach(function(client) {
      emailBody += formatClientForEmail(client);
    });
  } else {
    emailBody += "No clients specifically for " + todayDay + "\n\n";
  }

  // 5b) Then “other” clients
  if (otherClients.length > 0) {
    emailBody += "=== REST OF THE WEEK ===\n\n";
    otherClients.forEach(function(client) {
      emailBody += formatClientForEmail(client);
    });
  }

  // 6) Send
  MailApp.sendEmail({
    to: "rbarrio1@alumni.nd.edu",  // <-- change to your desired email
    subject: "Immediate Summary (" + todayDay + ")",
    body: emailBody
  });
}


/**
 * findLargestDateAmong(arrayOfStrings):
 * Among an array of multi-line text blocks (like client.followUps),
 * returns the single largest date found in ANY line. If none, returns null.
 */
function findLargestDateAmong(blocks) {
  if (!blocks || !blocks.length) return null;
  var bestDate = null;
  blocks.forEach(function(block) {
    var lines = block.split("\n");
    lines.forEach(function(line) {
      var dt = extractDateFromLine(line);
      if (dt && (!bestDate || dt > bestDate)) {
        bestDate = dt;
      }
    });
  });
  return bestDate;
}


/**
 * formatClientForEmail(client):
 * Returns a string in the exact format:
 *
 *   ClientName (Category)
 *   Previous: <line from pastWorks>
 *   Next: <line from followUps>
 *   In Progress: <columnLContent>
 *
 * We pick the single best “most recent” line from pastWorks, and from followUps.
 */
function formatClientForEmail(client) {
  var text = "";

  // 1) The name + category
  text += client.name;
  if (client.category) {
    text += " (" + client.category + ")";
  }
  text += "\n";

  // 2) Past Work: find the single line with the largest date
  var pastLine = findLargestDatedLine(client.pastWorks);
  text += "Previous: " + (pastLine || "N/A") + "\n";

  // 3) Future Work: single largest dated line
  var futureLine = findLargestDatedLine(client.followUps);
  text += "Next: " + (futureLine || "N/A") + "\n";

  // 4) In Progress (Column L)
  var inProg = (client.columnLContent || "").trim();
  if (inProg) {
    text += "In Progress: " + inProg + "\n";
  } else {
    text += "In Progress: (none)\n";
  }

  text += "\n";  // blank line after each client
  return text;
}


/**
 * findLargestDatedLine(arrayOfBlocks):
 * Goes through each block and each line, searching for the line
 * whose date is the largest. Returns just that line’s text.
 * If none is found, returns null.
 */
function findLargestDatedLine(blocks) {
  if (!blocks || !blocks.length) return null;
  var bestLine = null;
  var bestDate = null;

  blocks.forEach(function(block) {
    var lines = block.split("\n");
    lines.forEach(function(line) {
      var dt = extractDateFromLine(line);
      if (dt && (!bestDate || dt > bestDate)) {
        bestDate = dt;
        bestLine = line;
      }
    });
  });

  return bestLine;
}


/**
 * extractDateFromLine(line):
 * EXACT copy from your existing code that parses "MM/DD/YY" or "MM/DD/YYYY"
 * and returns a Date object. Returns null if not found / invalid.
 *
 * Make sure you have this in your code or else define it exactly as below.
 */
function extractDateFromLine(line) {
  if (!line) return null;
  var match = line.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (match) {
    var month = parseInt(match[1], 10) - 1;
    var day   = parseInt(match[2], 10);
    var year  = parseInt(match[3], 10);
    if (year < 100) {
      year = 2000 + year;
    }
    var dt = new Date(year, month, day);
    if (!isNaN(dt.getTime())) {
      return dt;
    }
  }
  return null;
}
/******************************************************************************************************
END SNIPPET
******************************************************************************************************/





////////////////////////////////////////////////////////////////////////////////////////////////
// HELPER FUNCTIONS - Copy/paste these below your getTopClients() or near the bottom of Code.gs
////////////////////////////////////////////////////////////////////////////////////////////////

/**
 * Among an array of multi-line strings, returns the single line containing the *most recent* date.
 * If none have a valid date, it returns null.
 */
function findLineWithLargestDate(textArray) {
  if (!textArray || !textArray.length) return null;

  var mostRecentLine = null;
  var mostRecentDate = null;

  textArray.forEach(function(blockOfLines) {
    var lines = blockOfLines.split("\n");
    lines.forEach(function(line) {
      var lineDate = extractDateFromLine(line);
      if (lineDate && (!mostRecentDate || lineDate > mostRecentDate)) {
        mostRecentDate = lineDate;
        mostRecentLine = line;
      }
    });
  });

  return mostRecentLine;
}

/**
 * Ensure a time-based trigger exists to send the JB chip summary at 8 AM daily.
 */
function ensureJbChipDailyTrigger() {
  var handler = 'sendJbChipTasksEmail';
  var hasTrigger = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === handler;
  });

  if (!hasTrigger) {
    ScriptApp.newTrigger(handler)
      .timeBased()
      .atHour(8)
      .everyDays(1)
      .create();
  }
}

function parseDateFromLine_(line) {
  if (!line) return null;
  var match = String(line).match(/\b(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/);
  if (!match) return null;

  var month = parseInt(match[1], 10) - 1;
  var day = parseInt(match[2], 10);
  var year = parseInt(match[3], 10);
  if (year < 100) {
    year += 2000;
  }

  var dt = new Date(year, month, day);
  return isNaN(dt.getTime()) ? null : dt;
}

function formatDateTaskLine_(line, tz) {
  var match = String(line).match(/\b(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/);
  if (!match) {
    return String(line || '').trim();
  }

  var date = parseDateFromLine_(line);
  var formattedDate = date
    ? Utilities.formatDate(date, tz, "MM/dd/yy")
    : match[0];

  var remainder = String(line).replace(match[0], '').trim().replace(/^[:\-–\s]+/, '');
  return remainder ? (formattedDate + ': ' + remainder) : formattedDate;
}

function extractHighlightedLines_(lines, tz) {
  var withDates = [];
  var withoutDates = [];

  (lines || []).forEach(function(rawLine) {
    var line = String(rawLine || '').trim();
    if (!line) return;

    var dt = parseDateFromLine_(line);
    if (dt) {
      withDates.push({ date: dt, text: formatDateTaskLine_(line, tz) });
    } else {
      withoutDates.push({ text: line });
    }
  });

  withDates.sort(function(a, b) { return a.date - b.date; });
  var ordered = withDates.map(function(item) { return item.text; })
    .concat(withoutDates.map(function(item) { return item.text; }));

  var start = Math.max(0, ordered.length - 3);
  return ordered.slice(start);
}

function collectJbChipClients_(targetDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  if (!sheet) {
    return { dateString: '', clients: [] };
  }

  var data = sheet.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var targetStr = Utilities.formatDate(targetDate, tz, "MM/dd/yy");
  var clients = {};

  data.forEach(function(row, idx) {
    if (idx === 0) return; // header

    var rawName = row[0];
    if (!rawName) return;

    var initials = row[15] ? row[15].toString().trim().toUpperCase() : '';
    var chipDateRaw = row[16];
    var chipDateStr = '';
    if (chipDateRaw) {
      var chipDate = (chipDateRaw instanceof Date) ? chipDateRaw : new Date(chipDateRaw);
      chipDateStr = isNaN(chipDate.getTime())
        ? chipDateRaw.toString()
        : Utilities.formatDate(chipDate, tz, "MM/dd/yy");
    }

    if (initials !== 'JB' || chipDateStr !== targetStr) {
      return;
    }

    var name = rawName.toString().replace(/\d+$/, '').trim();
    if (!clients[name]) {
      clients[name] = { name: name, pastWorks: [] };
    }

    var pastNote = row[2];
    if (pastNote) {
      String(pastNote).split('\n').forEach(function(line) {
        var t = line.trim();
        if (t) {
          clients[name].pastWorks.push(t);
        }
      });
    }
  });

  var sortedNames = Object.keys(clients).sort(function(a, b) {
    return a.localeCompare(b);
  });

  var resultClients = sortedNames.map(function(name) {
    var entry = clients[name];
    return {
      name: entry.name,
      highlights: extractHighlightedLines_(entry.pastWorks, tz)
    };
  });

  return { dateString: targetStr, clients: resultClients };
}

function sendJbChipTasksEmail() {
  var summary = collectJbChipClients_(new Date());
  var lines = [];

  lines.push('JB Tasks for ' + (summary.dateString || 'today') + ':');
  lines.push('');

  if (!summary.clients.length) {
    lines.push('No clients matched the JB chip for today.');
  } else {
    summary.clients.forEach(function(client) {
      lines.push(client.name);

      if (client.highlights.length === 0) {
        lines.push('  No highlighted tasks available.');
      } else {
        client.highlights.forEach(function(line) {
          lines.push('  ' + line);
        });
      }

      lines.push('');
    });
  }

  MailApp.sendEmail({
    to: 'rbarrios815@gmail.com,jbgreatfamily1@gmail.com',
    subject: 'JB Tasks for ' + (summary.dateString || 'today'),
    body: lines.join('\n')
  });

  ensureJbChipDailyTrigger();
}

/***********************************************************************
 SNIPPET #4: SERVER-SIDE TASKS INTEGRATION
***********************************************************************/

/**
 * Returns an array of objects representing tasks from "Tasks" spreadsheet.
 * Each object has { fullTask: "...", rowIdInTasks: N }.
 * Adjust the sheet name & columns to match your actual Tasks layout.
 */
function getTasksForDashboardIntegration() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Tasks"); // <-- rename if needed
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues(); // entire sheet
  var tasksArray = [];

  // Assuming header row in row 1
  // And the 'Task' text is in column A, for example
  // rowIdInTasks = the actual row index (1-based)
  for (var r = 2; r <= data.length; r++) {
    var taskString = data[r-1][0]; // Column A
    if (taskString) {
      tasksArray.push({
        fullTask: taskString,
        rowIdInTasks: r // remember which row
      });
    }
  }
  return tasksArray;
}

/**
 * Deletes the specified row from "Tasks" spreadsheet.
 */
function deleteTaskRow(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Tasks"); // rename if needed
  if (!sheet) throw new Error("Tasks sheet not found!");

  sheet.deleteRow(rowIndex);
  return true;
}

////////////////////////////////////////////////////////////
/// START OF SNIPPET: getClosestClients(namePart, topN) ///
////////////////////////////////////////////////////////////
function getClosestClients(namePart, topN) {
  // 1) Normalize the user input (the partial name we’re matching).
  const inputTokens = normalizeNameParts_(namePart);

  // 2) Read all client names from your dashboard (Column A).
  //    Adjust the sheet name & range references as needed for your app.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DASHBOARD 8.0"); 
  if (!sheet) {
    return [];
  }
  // Assuming Column A has client names starting at row 2:
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const range = sheet.getRange(2, 1, lastRow - 1, 1);
  const values = range.getValues(); // 2D array of [[clientName], [clientName], ...]

  // 3) For each client name in column A, compute a "match score".
  //    We'll track { name, score } in an array.
  let scoredClients = [];
  for (let i = 0; i < values.length; i++) {
    const rowName = (values[i][0] || "").toString().trim();
    if (!rowName) continue;

    const rowTokens = normalizeNameParts_(rowName);

    const score = computeMatchScore_(inputTokens, rowTokens);
    scoredClients.push({ name: rowName, score: score });
  }

  // 4) Sort the clients by descending score. If tie, by name ascending.
  scoredClients.sort(function(a, b) {
    if (b.score !== a.score) {
      return b.score - a.score; // higher score first
    }
    return a.name.localeCompare(b.name); // alphabetical if same score
  });

  // 5) Slice topN or top5
  const N = topN || 5;
  const topMatches = scoredClients.slice(0, N).map(obj => obj.name);

  return topMatches;
}

/**
 * A helper function that removes “couple”, “and”, “&”, punctuation, etc.,
 * then splits the string into individual words in lowercase.
 */
function normalizeNameParts_(raw) {
  if (!raw) return [];
  
  // Convert hyphens to spaces so “Marin‑Favela” becomes “Marin Favela”
  let cleaned = raw.replace(/-/g, " ");

  // Remove “couple”, “and”, “&”
  // Also remove other small filler words if you’d like
  const ignoreWords = ["couple", "and", "&"];
  ignoreWords.forEach(w => {
    const re = new RegExp("\\b" + w + "\\b", "gi");
    cleaned = cleaned.replace(re, "");
  });

  // Remove any punctuation except spaces
  cleaned = cleaned.replace(/[.,/#!$%^*;:{}=`~()]/g, "");

  // Trim extra spaces & convert to lowercase
  cleaned = cleaned.trim().toLowerCase();

  // Split into tokens by whitespace
  const tokens = cleaned.split(/\s+/).filter(Boolean);
  return tokens;
}

/**
 * A function that computes a match score, focusing heavily on last names if possible.
 * We do a simple approach:
 *   +1 point for each word in input that also appears in row
 *   +2 bonus if the last word in input matches the last word in row
 */
function computeMatchScore_(inputParts, rowParts) {
  if (!inputParts.length || !rowParts.length) {
    return 0;
  }

  let score = 0;
  // For each token in input, +1 if rowParts contains it
  inputParts.forEach(token => {
    if (rowParts.indexOf(token) !== -1) {
      score += 1;
    }
  });

  // Bonus for last name match if both sets have at least 1 token
  const inputLast = inputParts[inputParts.length - 1];
  const rowLast = rowParts[rowParts.length - 1];
  if (inputLast && rowLast && inputLast === rowLast) {
    score += 2;
  }

  return score;
}
//////////////////////////////////////////////////////////
/// END OF SNIPPET: getClosestClients(namePart, topN) ///
//////////////////////////////////////////////////////////


function similarityScore(str1, str2) {
  str1 = str1.toLowerCase().trim();
  str2 = str2.toLowerCase().trim();

  var distance = levenshteinDistance(str1, str2);
  var maxLen = Math.max(str1.length, str2.length);
  if (maxLen === 0) return 1;

  return 1 - (distance / maxLen); // normalized to [0,1], higher is better
}

function levenshteinDistance(a, b) {
  var matrix = [];

  for (var i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }
  for (var j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }

  for (var i = 1; i <= b.length; i++) {
    for (var j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j] + 1      // deletion
        );
      }
    }
  }

  return matrix[b.length][a.length];
}






/**
 * Updates Columns P (initials) and Q (date) for the LAST occurrence row of a client.
 * initials must be one of: RB, JB, QC.
 * dateInput can be "YYYY-MM-DD" (from <input type="date">) or "MM/DD/YY" or a Date.
 */
/**
 * Updates Columns P (initials) and Q (date) for the LAST occurrence row of a client.
 * initials must be one of: RB, JB, QC, or blank.
 * dateInput can be "YYYY-MM-DD" (from <input type="date">), "MM/DD/YY(YY)", a Date, or blank.
 */
function updateClientChip(clientName, initials, dateInput) {
  var allowed = Object.create(null);
  ['RB','JB','TEAM','QC','BDAY',''].forEach(function(k){ allowed[k] = true; });

  var initialsTrim = (initials || '').toUpperCase().trim();
  if (!allowed[initialsTrim]) {
    throw new Error('Initials must be RB, JB, TEAM, QC, BDAY, or blank.');
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data  = sheet.getDataRange().getValues();
  var lastRowToUpdate = -1;

  // Find last row for this client (strip trailing digits like "Smith 2")
  for (var i = 1; i < data.length; i++) {
    var nameCell = data[i][0];
    if (!nameCell) continue;
    var cleaned = nameCell.toString().replace(/\d+$/, '').trim().toLowerCase();
    if (cleaned === clientName.trim().toLowerCase()) {
      lastRowToUpdate = i;
    }
  }
  if (lastRowToUpdate === -1) {
    throw new Error('Client not found: ' + clientName);
  }

  // P=16, Q=17 (1-based)
  var initialsRange = sheet.getRange(lastRowToUpdate + 1, 16);
  var dateRange     = sheet.getRange(lastRowToUpdate + 1, 17);

  initialsRange.setValue(initialsTrim);

  var dateToSet = '';
  if (dateInput) {
    if (dateInput instanceof Date) {
      dateToSet = dateInput;
    } else {
      var s = dateInput.toString().trim();

      // YYYY-MM-DD
      var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (iso) {
        dateToSet = new Date(iso[1], parseInt(iso[2],10)-1, iso[3]);
      } else {
        // Try native (MM/DD/YYYY) or similar
        var d = new Date(s);
        if (!isNaN(d.getTime())) {
          dateToSet = d;
        } else {
          // MM/DD/YY
          var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
          if (m) {
            var year = parseInt(m[3], 10);
            if (year < 100) year = 2000 + year;
            dateToSet = new Date(year, parseInt(m[1],10)-1, parseInt(m[2],10));
          } else {
            throw new Error('Invalid date format. Use YYYY-MM-DD or MM/DD/YY.');
          }
        }
      }
    }
  }

  if (dateToSet) dateRange.setValue(dateToSet);
  else           dateRange.clearContent();

  return true;
}

function updateInProgress(clientName, text) {
  var sh = SpreadsheetApp.getActive().getSheetByName('DASHBOARD 8.0');
  var values = sh.getDataRange().getValues();
  // Find the FIRST row whose Column A (name, with trailing digits stripped) matches
  for (var r = 1; r < values.length; r++) {
    var raw = values[r][0];
    if (!raw) continue;
    var base = String(raw).replace(/\d+$/, '').trim();
    if (base === clientName) {
      sh.getRange(r+1, 12).setValue(text); // Column L = 12
      return true;
    }
  }
  throw new Error('Client not found for Column L update: ' + clientName);
}

/*******************************
 * RECENT CLIENTS (per-user)
 * - Stores in UserProperties as JSON
 * - Each entry: { name, tsISO, category, columnL, chipInitials, chipDate }
 *******************************/

// Utility: get latest sheet details for a client (last occurrence row)
function _getLatestClientRowDetails_(clientName) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  var data = sh.getDataRange().getValues();
  var lastIdx = -1;

  for (var i = 1; i < data.length; i++) {
    var a = data[i][0];
    if (!a) continue;
    var base = String(a).replace(/\d+$/, '').trim().toLowerCase();
    if (base === clientName.trim().toLowerCase()) lastIdx = i;
  }
  if (lastIdx < 0) return null;

  var row = data[lastIdx];
  var tz  = Session.getScriptTimeZone();

  // columns: F=5, L=11, P=15, Q=16 (0-based indexing from getValues)
  var category     = row[5]  ? String(row[5]).trim()  : '';
  var columnL      = row[11] == null ? '' : String(row[11]);
  var chipInitials = row[15] ? String(row[15]).trim() : '';

  var chipDate = '';
  if (row[16]) {
    var d = (row[16] instanceof Date) ? row[16] : new Date(row[16]);
    chipDate = isNaN(d.getTime()) ? String(row[16]) : Utilities.formatDate(d, tz, "MM/dd/yy");
  }

  return {
    category: category,
    columnL: columnL,
    chipInitials: chipInitials,
    chipDate: chipDate
  };
}

// Add/refresh a client at the top of the recent list, keep unique, cap at N.
function logRecentClient(clientName, cap) {
  if (!clientName) return false;

  var details = _getLatestClientRowDetails_(clientName) || {
    category: '',
    columnL: '',
    chipInitials: '',
    chipDate: ''
  };

  var up = PropertiesService.getUserProperties();
  var key = 'recentClients';
  var capSize = Math.max(1, cap || 3);

  var existing = [];
  try {
    existing = JSON.parse(up.getProperty(key) || '[]');
  } catch (e) {
    existing = [];
  }

  // remove any existing entries with same name (case-insensitive compare)
  var lname = clientName.trim().toLowerCase();
  existing = existing.filter(function(it) {
    return (it && it.name && it.name.trim().toLowerCase() !== lname);
  });

  // unshift new
  existing.unshift({
    name: clientName,
    tsISO: new Date().toISOString(),
    category: details.category,
    columnL: details.columnL,
    chipInitials: details.chipInitials,
    chipDate: details.chipDate
  });

  // cap
  existing = existing.slice(0, capSize);

  up.setProperty(key, JSON.stringify(existing));
  return true;
}

// Return recent list (optionally refresh metadata so it never goes stale)
function getRecentClients(limit, refreshMetadata) {
  var up = PropertiesService.getUserProperties();
  var key = 'recentClients';
  var list = [];
  try {
    list = JSON.parse(up.getProperty(key) || '[]');
  } catch (e) {
    list = [];
  }

  // Optionally re-pull current sheet metadata for each client (slower, but up-to-date)
  if (refreshMetadata) {
    list = list.map(function(entry) {
      var d = _getLatestClientRowDetails_(entry.name);
      if (d) {
        entry.category = d.category;
        entry.columnL = d.columnL;
        entry.chipInitials = d.chipInitials;
        entry.chipDate = d.chipDate;
      }
      return entry;
    });
    up.setProperty(key, JSON.stringify(list)); // keep cache fresh
  }

  var lim = Math.max(1, limit || 3);
  return list.slice(0, lim);
}

// Clear all recents for the current user (handy for debugging)
function clearRecentClients() {
  PropertiesService.getUserProperties().deleteProperty('recentClients');
  return true;
}

/** ============ NOTES INBOX (A=Note, B=Assigned Client, C=Timestamp) ============ */
const NOTES_SHEET = 'NOTES INBOX';

function ensureNotesInbox_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(NOTES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(NOTES_SHEET);
    sh.getRange(1,1,1,3).setValues([['NOTE','CLIENT NAME (ASSIGNED)','TIMESTAMP']]);
  }
  return sh;
}

function inboxAddNote(rawNote, assignedClient) {
  // === CONFIG: who receives the "ADD" email ===
  var EMAIL_TO = 'rbarrio1@alumni.nd.edu';  // change if needed

  if (!rawNote || !String(rawNote).trim()) {
    throw new Error('Note is empty.');
  }

  // Ensure the NOTES INBOX sheet exists and append the note + timestamp
  const sh = ensureNotesInbox_(); // uses your existing helper
  const row = sh.getLastRow() + 1;

  var cleanNote = String(rawNote).trim();
  sh.getRange(row, 1).setValue(cleanNote);   // Col A: NOTE
  sh.getRange(row, 3).setValue(new Date());  // Col C: TIMESTAMP

  // Optionally assign to a client immediately (Column B)
  if (assignedClient && String(assignedClient).trim()) {
    var resolved = getCanonicalClientName_(assignedClient) || String(assignedClient).trim();
    sh.getRange(row, 2).setValue(resolved);
  }

  // Build a helpful subject and body for the email
  var tz = Session.getScriptTimeZone();
  var nowStr = Utilities.formatDate(new Date(), tz, 'MM/dd/yy h:mma');

  // Use first ~60 chars of the note in subject (sanitized)
  var snippet = cleanNote.replace(/\s+/g, ' ').slice(0, 60);
  if (cleanNote.length > 60) snippet += '…';

  var subject = 'DASHBOARD UPDATE: ' + snippet;
  var body =
    'DASHBOARD UPDATE: ' + cleanNote;


  // Send the email (don’t block saving on email errors)
  try {
    MailApp.sendEmail({ to: EMAIL_TO, subject: subject, body: body });
  } catch (e) {
    // Log but continue (the note is already saved)
    Logger.log('Email send failed: ' + e);
  }

  return { ok: true, row: row };
}


function inboxGetRecent(limit) {
  const sh = ensureNotesInbox_();
  const last = sh.getLastRow();
  if (last < 2) return { recent: [], raw: [] };

  const all = sh.getRange(2,1,last-1,3).getValues().map((r,i)=>({
    row: i+2, note: String(r[0]||'').trim(),
    assigned: String(r[1]||'').trim(),
    ts: r[2] ? new Date(r[2]).getTime() : 0
  }));

  const tz = Session.getScriptTimeZone();
  const sorted = [...all]
    .filter(x => x.ts) // only rows with a timestamp in Column C
    .sort((a,b)=>b.ts-a.ts)
    .slice(0, limit||5)
    .map(x=>({
    row: x.row,
    note: x.note,
    assigned: x.assigned,
    timestamp: x.ts ? Utilities.formatDate(new Date(x.ts), tz, 'MM/dd/yy h:mma') : ''
  }));

  // Pull chip data (Columns P/Q) for any assigned clients so the UI can render chips
  var assignedNames = all
    .map(function(entry){ return entry.assigned; })
    .filter(function(name){ return name; });
  var chipMap = {};
  try {
    chipMap = getChipStateForClients(assignedNames);
  } catch (e) {
    chipMap = {};
  }

  return { recent: sorted, raw: all, chipMap: chipMap };
}

function inboxUpdateNote(row, newNote){
  if (!row || row < 2) {
    throw new Error('A valid row number is required.');
  }

  const note = String(newNote || '').trim();
  if (!note) {
    throw new Error('Note cannot be empty.');
  }

  const sh = ensureNotesInbox_();
  const hasTimestamp = sh.getRange(row, 3).getValue();
  if (!hasTimestamp) {
    throw new Error('Cannot edit this note because Column C is missing a date.');
  }

  const oldNote = String(sh.getRange(row, 1).getValue() || '').trim();
  const assignedClient = String(sh.getRange(row, 2).getValue() || '').trim();
  sh.getRange(row, 1).setValue(note);

  let synced = true;
  if (assignedClient) {
    synced = syncInboxNoteToDashboard_(assignedClient, oldNote, note, hasTimestamp);
  }

  if (!synced) {
    throw new Error('EDIT FAILED ON OTHER VIEW');
  }

  return { ok: true, row: row, note: note, synced: synced };
}

function syncInboxNoteToDashboard_(clientName, oldNote, newNote, timestamp) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();
  const target = normalizeClientName_(clientName);
  const dateStr = timestamp
    ? Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'MM/dd/yy')
    : '';

  for (let i = 1; i < data.length; i++) {
    const rowName = normalizeClientName_(data[i][0]);
    if (rowName !== target) continue;

    const pastWork = data[i][2];
    const lines = String(pastWork || '').split('\n');
    const idx = findNoteLineIndex_(lines, oldNote, dateStr);
    if (idx === -1) continue;

    const prefix = extractDatePrefix_(lines[idx]);
    lines[idx] = (prefix ? prefix : '') + newNote;
    sheet.getRange(i + 1, 3).setValue(lines.join('\n'));
    return true;
  }

  return false;
}

function findNoteLineIndex_(lines, oldNote, dateStr) {
  let fallback = -1;
  const cleanOld = String(oldNote || '').trim();

  for (let i = 0; i < lines.length; i++) {
    const line = String(lines[i] || '');
    const noteText = line.replace(/^\s*\d{1,2}\/\d{1,2}\/\d{2,4}\s*:\s*/, '').trim();
    if (noteText !== cleanOld) continue;

    const matchesDate = dateStr && line.trim().startsWith(dateStr);
    if (matchesDate) return i;
    if (fallback === -1) fallback = i;
  }

  return fallback;
}

function extractDatePrefix_(line) {
  const m = String(line || '').match(/^\s*\d{1,2}\/\d{1,2}\/\d{2,4}\s*:\s*/);
  return m ? m[0] : '';
}

function normalizeClientName_(name) {
  return String(name || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\d+$/, '')
    .trim()
    .toLowerCase();
}

function getCanonicalClientName_(typed) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD 8.0');
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const vals = sheet.getRange(2,1,last-1,1).getValues().flat();
  const needle = String(typed||'').replace(/\d+$/,'').trim().toLowerCase();
  for (const v of vals) {
    const raw = String(v||'').trim();
    const base = raw.replace(/\d+$/,'').trim();
    if (base.toLowerCase() === needle) return base;
    if (raw.toLowerCase()  === needle) return raw;
  }
  return null;
}

function inboxAssignToClient(row, clientTypedName) {
  const sh = ensureNotesInbox_();
  const note = String(sh.getRange(row, 1).getValue() || '').trim();
  if (!note) throw new Error('No note in Column A for row ' + row + '.');

  const canonical = getCanonicalClientName_(clientTypedName);
  if (!canonical) throw new Error('Client not found on DASHBOARD 8.0: ' + clientTypedName);

  sh.getRange(row, 2).setValue(canonical);      // Column B = submitted/assigned
  const ok = updatePastWork(canonical, note);   // your existing submit into Client Details
  return { ok, client: canonical, row };
}

function getClientNamesSimple() {
  return getClientNames(); // your existing list (minus categories)
}


/**
 * Returns a map of normalized client name -> { initials, dateMMDDYY }
 * Reads Column A (name), Column P (initials), Column Q (date).
 */
function getChipStateForClients(clientNames) {
  var sh = SpreadsheetApp.getActive().getSheetByName('DASHBOARD 8.0');
  if (!sh) throw new Error("Sheet 'DASHBOARD 8.0' not found");
  var last = sh.getLastRow();
  if (last < 2) return {};

  function normName(s) {
    return String(s || '')
      .replace(/\u00A0/g, ' ')
      .replace(/\s+-\s+.*$/, '')
      .replace(/\d+$/, '')
      .trim()
      .toLowerCase();
  }

  var tz = Session.getScriptTimeZone();
  var rng = sh.getRange(2, 1, last - 1, 17).getValues(); // A..Q
  var map = {};
  for (var i = 0; i < rng.length; i++) {
    var row = rng[i];
    var a = row[0];           // name
    if (!a) continue;
    var k = normName(a);
    var initials = (row[15] || '').toString().toUpperCase(); // P
    var q        = row[16];                                   // Q
    var mmddyy   = '';
    if (q) {
      var d = (q instanceof Date) ? q : new Date(q);
      if (!isNaN(d)) {
        mmddyy = Utilities.formatDate(d, tz, "MM/dd/yy");
      } else {
        mmddyy = String(q);
      }
    }
    map[k] = { initials: initials, dateMMDDYY: mmddyy };
  }

  if (Array.isArray(clientNames) && clientNames.length) {
    var out = {};
    for (var j = 0; j < clientNames.length; j++) {
      var kk = normName(clientNames[j]);
      if (map.hasOwnProperty(kk)) out[kk] = map[kk];
    }
    return out;
  }
  return map;
}






/** === Siri → Google Sheets Bridge: Column A note, Column C timestamp ===
 * Appends note in Col A, leaves Col B blank, puts timestamp in Col C.
 */
const SHEET_ID = "1rzejdmR0hatqESPp9MroCwT229QGM0oB2G9mELaL4Ps";
const TARGET_SHEET = "NOTES INBOX";
const COL_NOTE = 1; // A
const COL_BLANK = 2; // B
const COL_TIME = 3; // C

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return _json({ ok: false, error: "No postData received" });
    }
    const body = JSON.parse(e.postData.contents);
    const note = (body && body.note != null) ? String(body.note).trim() : "";
    if (!note) return _json({ ok: false, error: "Missing 'note' value" });

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(TARGET_SHEET);
    if (!sh) return _json({ ok: false, error: `Sheet not found: ${TARGET_SHEET}` });

    const nextRow = sh.getLastRow() + 1;
    const now = new Date();
    sh.getRange(nextRow, COL_NOTE, 1, 3).setValues([[note, "", now]]);

    return _json({
      ok: true,
      addedTo: `${TARGET_SHEET}!A${nextRow}:C${nextRow}`,
      note,
      timestampISO: now.toISOString()
    });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}