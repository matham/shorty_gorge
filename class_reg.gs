// updateFormClassList runs daily
// parseSelectedClasses runs From spreadsheet - On form submit

// constants such as sheet names, column names etc
const scheduleSheetName = "Class schedule";
const responseSheetName = "Registrations via Google Form";
const scheduleStartDateColName = "Start Date";
const scheduleListedColName = "Listed";
const regEntryTitleName = "Which series would you like to register for?";
const classColInsertionAfter = "Macro processed";
const classParsedSignalCol = "Macro processed";

const scheduleColNamesMap = new Map([
  ["Class Name", null],
  ["Unique ID", null],
  ["Location", null],
  ["Price", null],
  ["Door Price", null],
  ["Dates", null],
  ["Form Notes", null],
  ["Description", null],
  ["Requirements", null],
  ["Length", null],
])


// updates the form with current active classes
function updateFormClassList() {
  const lock = LockService.getPublicLock();
  lock.waitLock(30000);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formUrl = ss.getSheetByName(responseSheetName).getFormUrl();
  const form = FormApp.openByUrl(formUrl);

  let items = form.getItems();
  let classes = getClasses();

  // get the entry item
  let entry = null;
  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle() === regEntryTitleName)
    {
      entry = items[i];
      break;
    }
  }
  if (entry === null) {
    throw new RangeError("Could not find registration entry");
  }

  // add current classes
  let description = "Select your class(es) below\n";
  let entries = [];
  for (i = 0; i < classes.length; i++) {
    let name = classes[i][scheduleColNamesMap.get("Class Name")] + " - " + classes[i][scheduleColNamesMap.get("Unique ID")];

    description += "\n\n-------  " + name + "  -------\n\n";
    description += classes[i][scheduleColNamesMap.get("Description")] + "\n\n";
    if (classes[i][scheduleColNamesMap.get("Form Notes")]) {
      description += classes[i][scheduleColNamesMap.get("Form Notes")] + "\n\n";
    }

    description += "-- Prerequisites: " + classes[i][scheduleColNamesMap.get("Requirements")] + "\n";
    description += "-- Length: " + classes[i][scheduleColNamesMap.get("Length")] + "\n";
    description += "-- Dates: " + classes[i][scheduleColNamesMap.get("Dates")] + "\n\n";

    description += "-- Room: " + classes[i][scheduleColNamesMap.get("Location")] + "\n";
    description += "-- Price: $" + classes[i][scheduleColNamesMap.get("Price")] + " for the full series ($"
      + classes[i][scheduleColNamesMap.get("Door Price")] + " at the door)\n";

    entries.push(name);
  }
  
  entry.setHelpText(description);
  entry.asCheckboxItem().setChoiceValues(entries);

  lock.releaseLock();
}


// gets all the current classes
function getClasses() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let scheduleSheet = ss.getSheetByName(scheduleSheetName);

  const lastRow = scheduleSheet.getLastRow();
  const lastColumn = scheduleSheet.getLastColumn();
  const scheduleValues = scheduleSheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const dateCol = scheduleValues[0].indexOf(scheduleStartDateColName);
  const listedCol = scheduleValues[0].indexOf(scheduleListedColName);

  for (const key of scheduleColNamesMap.keys()) {
    scheduleColNamesMap.set(key, scheduleValues[0].indexOf(key));
  }

  // current date
  let now = new Date();
  now.setHours(0,0,0,0);
  let classes = [];

  // skip header row
  for (let i = 1; i < scheduleValues.length; i++) {
    // end of content
    if (scheduleValues[i][0] === "") {
      break;
    }

    if (scheduleValues[i][dateCol] >= now && scheduleValues[i][listedCol] === "Yes") {
      classes.push(scheduleValues[i]);
    }
  }

  return classes;
}


// displays the emails and names for all rows currently visible in format you can paste into email "to" field
function displayEmailListForEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

  const nameCol = values[0].indexOf("Name");
  const emailCol = values[0].indexOf("Email Address");

  if (!nameCol || !emailCol) {
    throw new RangeError("Did not find the name or email column by name");
  }

  const emailSet = new Set();
  let items = "";
  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]){
      break;
    }

    if (!sheet.isRowHiddenByFilter(i) && !emailSet.has(values[i][emailCol])) {
      items += '"' + values[i][nameCol] + '" <' + values[i][emailCol] + '>, ';
      emailSet.add(values[i][emailCol]);
    }
  }

  let htmlOutput = HtmlService
      .createHtmlOutput('<p>' + escapeText(items) + '</p>')
      .setTitle('Copy filtered email list');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}


// displays the emails and names for people who want to be registered
function displayEmailListForMailchimp() {
  const lock = LockService.getPublicLock();
  lock.waitLock(30000);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

  const nameCol = values[0].indexOf("Name");
  const emailCol = values[0].indexOf("Email Address");
  const subCol = values[0].indexOf("Subscribe to our mailing list (emails not shared with anyone else)");
  const alreadySubCol = values[0].indexOf("Subscribed?");
  
  if (!nameCol || !emailCol || !subCol || !alreadySubCol) {
    throw new RangeError("Did not find the name, email, subscribe, or already subscribed column by name");
  }

  const emailSet = new Set();
  let items = "Email Address, Name<br/>";
  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]){
      break;
    }

    if (values[i][subCol] === "Yes" && values[i][alreadySubCol] !== "Yes") {
      sheet.getRange(i + 1, alreadySubCol + 1).setValue("Yes");
      if (!emailSet.has(values[i][emailCol])) {
        items += escapeText(values[i][emailCol] + ', ' + values[i][nameCol]) + '<br/>';
        emailSet.add(values[i][emailCol]);
      }
    }
  }
    
  let htmlOutput = HtmlService
      .createHtmlOutput('<p>' + items + '</p>')
      .setTitle('Copy Mailchimp email list');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}


function escapeText(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}


function parseSelectedClasses() {
  const lock = LockService.getPublicLock();
  lock.waitLock(30000);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(responseSheetName);

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

  const selectionCol = values[0].indexOf(regEntryTitleName);
  let insertionCol = values[0].indexOf(classColInsertionAfter);
  const processedCol = values[0].indexOf(classParsedSignalCol);

  if (!selectionCol || !insertionCol || !processedCol) {
    throw new RangeError("Did not find the required columns by name");
  }

  const classColMap = new Map();

  // for each row parse registered classes
  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]){
      break;
    }
    // already processed?
    if (values[i][processedCol] === "Yes") {
      continue;
    }

    // get list of registered classes for this row
    const names = values[i][selectionCol].split(",")
    const usedNames = new Map();
    for (let name of names) {
      name = name.trim().split(' - ');
      const iterName = name[name.length - 1];
      name = name.slice(0, -1).join(' - ');

      if (!name || !iterName) {
        throw new RangeError("Unable to parse class name from registration selection");
      }

      // first time we see this class add it
      if (!usedNames.has(name)) {
        usedNames.set(name, []);
      }
      usedNames.get(name).push(iterName);
    }

    // add parsed classes to individual columns
    for (let [name, iterNames] of usedNames) {
      // class doesn't have col yet
      if (!classColMap.has(name)) {
        const existCol = values[0].indexOf(name);
        // col we store in map is zero based. We write as one based
        if (existCol >= 0) {
          classColMap.set(name, existCol);
        } else {
          // add new col and set header
          sheet.insertColumnAfter(insertionCol + 1);
          sheet.getRange(1, insertionCol + 2).setValue(name);

          classColMap.set(name, insertionCol + 1);
          insertionCol++;
        }
      }

      const cell = sheet.getRange(i + 1, classColMap.get(name) + 1);
      cell.setValue(iterNames.join(", "));
      cell.setNumberFormat("@");
    }
    
    sheet.getRange(i + 1, processedCol + 1).setValue("Yes");
  }

  lock.releaseLock();
}
