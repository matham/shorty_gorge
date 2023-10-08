const toProcessLabelName = 'Payment/PaymentToBeProcessed';
const processedLabelName = 'Payment/PaymentProcessed';
const eChargesSheetName = 'online charges';
const venmoSheetName = 'venmo acct';
const paySourceForCC = 'Mario CC';


function parseEmails() {
  const toProcessLabel = GmailApp.getUserLabelByName(toProcessLabelName);
  const processedLabel = GmailApp.getUserLabelByName(processedLabelName);

  const lock = LockService.getPublicLock();
  lock.waitLock(30000);

  const threadsToProcess = toProcessLabel.getThreads(0, 30);
  if (!threadsToProcess.length) {
    return;
  }

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {'ss': ss, 'eCharges': null, 'venmo': null};

  for (let i = 0; i < threadsToProcess.length; i++) {
    const thread = threadsToProcess[i];
    if (parseEmail(thread, sheets)) {
      thread.addLabel(processedLabel);
      thread.removeLabel(toProcessLabel);
      thread.markRead();
    }
  }

  lock.releaseLock();
}

function parseEmail(thread, sheets) {
  const subject = thread.getFirstMessageSubject();
  const messages = thread.getMessages();

  if (messages.length === 0) {
    throw new TypeError('No message in the email');
  }

  const msg = messages[0];
  const msgFrom = msg.getFrom().toLowerCase();

  if (msgFrom.includes('hover')) {
    return parseHover(msg, subject, sheets);
  } else if (msgFrom.includes('mailchimp')) {
    return parseMailchimp(msg, subject, sheets);
  } else if (msgFrom.includes('mariomtfg@gmail.com') || msgFrom.includes('venmo')) {
    return parseVenmo(msg, subject, sheets);
  }
  return false;
}

function parseHover(msg, subject, sheets) {
  const body = msg.getPlainBody();
  const date = body.match(/Order Date: *([0-9]+)-([0-9]+)-([0-9]+)/i);
  const note = body.match(/Order ID: *([\w\-]+)/i);
  const amt = body.match(/Order Total: *\$([0-9.]+)/i);

  if (!date || !note || !amt) {
    Logger.log('Could not match email with subject and body:');
    Logger.log([date, note, amt]);
    Logger.log(subject);
    Logger.log(body);
    return false;
  }

  const dateStr = date[3] + '/' + date[2] + '/' + date[1];

  appendEChargesRow([dateStr, 'Hover', amt[1], note[1], paySourceForCC], sheets);
  return true;
}

function parseVenmo(msg, subject, sheets) {
  const body = msg.getPlainBody();

  if (subject.includes('commented on a payment') || subject.includes('statement is ready')) {
    return true;
  }

  let row = null;
  if (body.match(/paid\s+you/i)) {
    row = parseVenmoPaidMe(body);
  } else if (body.match(/You\s+\<.+user_id.+\>\s+charged/i)) {
    row = parseVenmoPaidCharge(body);
  } else if (body.match(/You\s+\<.+user_id.+\>\s+paid/i)) {
    row = parseVenmoWePaid(body);
  } else if (body.match(/Standard +Transfer +Initiated/i)) {
    row = parseVenmoTransfer(body);
  } else {
    Logger.log('Could not match email with subject and body:');
    Logger.log(subject);
    Logger.log(body);
    return false;
  }

  if (!row) {
    Logger.log('Could not match email with subject and body:');
    Logger.log(subject);
    Logger.log(body);
    return false;
  }

  let sheet = sheets.venmo;
  if (!sheet) {
    sheets.venmo = sheets.ss.getSheetByName(venmoSheetName);
    sheet = sheets.venmo;
  }
  //Logger.log(row);
  sheet.appendRow(row);
  return true;
}

function cleanVenmoDate(date) {
  if (date.endsWith('PDT') || date.endsWith('PST')) {
    date = date.substring(0, date.length - 3).trim();
  }
  return date
}

function parseVenmoPaidMe(body) {
  const re = /\<.+?venmo.+?\>\s*(.+?)\s*\<.+?venmo.+?user_id\=([\d]+).+\>\s*paid\s+You\s*\<.+?venmo.+?\>\s*([^]*?)\s*Transfer Date and Amount:\s*([\w ,]+)·[^]+?\$([\d.]+)\s*Fee\s*\-\s*\$([\d.]+)\s*\+\s*\$([\d.]+)/i;
  const match = body.match(re);
  if (!match) {
    return false;
  }

  const row = [
    cleanVenmoDate(match[4].trim()),
    match[1].trim() + "<https://venmo.com/code?user_id=" + match[2].trim() + ">",
    match[7].trim(),
    "",
    match[3].trim(),
    match[5].trim(),
    match[6].trim()
  ]
  return row;
}

function parseVenmoPaidCharge(body) {
  
}

function parseVenmoWePaid(body) {
  const re = /You\s*\<.+?venmo.+?\>\s*paid\s*(.+?)\s*\<.+?venmo.+?user_id\=([\d]+).+\>\s*([^]*?)\s*Transfer Date and Amount:\s*([\w ,]+)·.+?\$([\d.]+)/i;
  const match = body.match(re);
  if (!match) {
    return false;
  }

  const row = [
    cleanVenmoDate(match[4].trim()),
    match[1].trim() + "<https://venmo.com/code?user_id=" + match[2].trim() + ">",
    "",
    match[5].trim(),
    match[3].trim()
  ]
  return row;
}

function parseVenmoTransfer(body) {
  
}

function parseMailchimp(msg, subject, sheets) {
  const body = msg.getPlainBody();
  const date = body.match(/(?:\r|\n)Processed on +(.+?) +New York\. *(?:\r|\n)/i);
  const note = body.match(/(?:\r|\n)(?:Order|Invoice) +([\w\-]+)(?:\r|\n)/i);
  const amt = body.match(/(?:\r|\n)Total +\$([0-9.]+)(?:\r|\n)/i);

  if (!date || !note || !amt) {
    Logger.log('Could not match email with subject and body:');
    Logger.log([date, note, amt]);
    Logger.log(subject);
    Logger.log(body);
    return false;
  }

  appendEChargesRow([date[1], 'Mailchimp', amt[1], note[1], paySourceForCC], sheets);
  return true;
}

function appendEChargesRow(row, sheets) {
  let sheet = sheets.eCharges;
  if (!sheet) {
    sheets.eCharges = sheets.ss.getSheetByName(eChargesSheetName);
    sheet = sheets.eCharges;
  }

  sheet.appendRow(row);
}
