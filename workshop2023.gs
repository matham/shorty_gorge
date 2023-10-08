function fillNew() {
  var spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.getRange('F2:H2').autoFill(spreadsheet.getRange('F2:H'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  var dataRange = spreadsheet.getRange('I2:I');
  var data = dataRange.getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === "") {
      spreadsheet.getRange("I" + (i + 2).toString()).setValue(spreadsheet.getRange("H" + (i + 2).toString()).getValue());
    }
  }

  var dataRange = spreadsheet.getRange('J2:J');
  var data = dataRange.getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === "") {
      spreadsheet.getRange("J" + (i + 2).toString()).setValue("Not paid");
    }
  }

  var dataRange = spreadsheet.getRange('M2:M');
  var data = dataRange.getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === "") {
      spreadsheet.getRange("M" + (i + 2).toString()).setValue("No");
    }
  }
};


function escapeText(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
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

    if (!sheet.isRowHiddenByFilter(i + 1) && !emailSet.has(values[i][emailCol])) {
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
