function tagEmails() {
  var emailDomain = "email.cynthia.re"; // The domain part of the email addresses
  var labelPrefix = "EvilEmail/"; // Prefix for the labels
  var tbpLabelName = labelPrefix + "TBP"; // Label for "To be processed"
  
  var tbpLabel = GmailApp.getUserLabelByName(tbpLabelName);
  
  var mapAddrToTag = {};
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var cells = sheet.getRange('A2:D1000');
  var cellRows = cells.getValues();
  for (var i in cellRows) {
    var key = String(cellRows[i][0]).toLowerCase();
    var keyName = String(cellRows[i][1]);
    var isActive = String(cellRows[i][2]).toLowerCase() === "yes";
    if (!isActive) {
      continue;
    }
    mapAddrToTag[key + '@' + emailDomain] = keyName;
  }
  
  console.log(mapAddrToTag);
  
  var threads = tbpLabel.getThreads();
  for (var i in threads) {
    var messages = threads[i].getMessages();

    var toAddr = messages[0].getHeader("X-Gm-Original-To").toLowerCase(); 
    
    if (!(toAddr in mapAddrToTag)) {
      console.log("noop: \"" + threads[i].getFirstMessageSubject() + "\"");
      continue;
    }
    
    var tagName = labelPrefix + mapAddrToTag[toAddr];
    var tag = GmailApp.getUserLabelByName(tagName);
    if (tag == null) {
      tag = GmailApp.createLabel(tagName);
    }
    tag.addToThread(threads[i]);
    
    tbpLabel.removeFromThread(threads[i]);
  }
}
