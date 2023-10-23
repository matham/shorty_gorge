const toProcessVenmoLabelStr = "RawVenmo";
const marioVenmoLabelStr = "MarioVenmo";
const shortyVenmoLabelStr = "TSGVenmo";
const fwdEmail = "theshortygorge@gmail.com";
const shortyVenmoUserId = "3641914934953114680";
const scriptTimeQuotaMS = 6 * 60 * 1000;

function filterVenmoEmails() {
  const tStart = Date.now();
  if (MailApp.getRemainingDailyQuota() < 3) {
    return;
  }
  
  const toProcessLabel = GmailApp.getUserLabelByName(toProcessVenmoLabelStr);
  const marioLabel = GmailApp.getUserLabelByName(marioVenmoLabelStr);
  const shortyLabel = GmailApp.getUserLabelByName(shortyVenmoLabelStr);


  const lock = LockService.getPublicLock();
  lock.waitLock(30000);

  const threadsToProcess = toProcessLabel.getThreads(0, 30);

  for (let i = 0; i < threadsToProcess.length && Date.now() - tStart < 0.5 * scriptTimeQuotaMS; i++) {
    const thread = threadsToProcess[i];
    const msg = thread.getMessages()[0];

    thread.removeLabel(toProcessLabel);

    if (msg.getBody().includes(shortyVenmoUserId)) {
      thread.addLabel(shortyLabel);
      msg.forward(fwdEmail);
      thread.markRead();
    } else {
      thread.addLabel(marioLabel);
    }
  }

  lock.releaseLock();
}
