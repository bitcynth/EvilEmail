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
    
    var tagName = String(mapAddrToTag[toAddr]);
    var prefixedTag = true;
    if (tagName.charAt(0) === '/') {
      prefixedTag = false;
      tagName = tagName.substr(1);
    } else {
      tagName = labelPrefix + tagName;
    }

    var tag = GmailApp.getUserLabelByName(tagName);
    if (tag == null && prefixedTag) {
      tag = GmailApp.createLabel(tagName);
    }

    // If tag still doesn't exist, noop
    if (tag == null) {
      console.log("noop (no tag): \"" + threads[i].getFirstMessageSubject() + "\"");
      continue;
    }

    tag.addToThread(threads[i]);
    
    tbpLabel.removeFromThread(threads[i]);
  }
}
