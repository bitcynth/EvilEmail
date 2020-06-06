function tagEmails() {
  var emailDomain = "email.cynthia.re";
  
  var threads = GmailApp.getInboxThreads();

  // Tag for "To be processed"
  var tbpLabel = GmailApp.getUserLabelByName("EvilEmail/TBP");
  
  var mapAddrToTag = {};
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var cells = sheet.getRange('A2:D100');
  var cellRows = cells.getValues();
  for (var i in cellRows) {
    if (cellRows[i][2].toLowerCase() != "yes") {
      continue;
    }
    mapAddrToTag[cellRows[i][0] + '@' + emailDomain] = cellRows[i][1];
  }
  
  console.log(mapAddrToTag);
  
  var threads = tbpLabel.getThreads();
  for (var i in threads) {
    var messages = threads[i].getMessages();
    var toHeader = messages[0].getTo();
    if (!(toHeader.toLowerCase() in mapAddrToTag)) {
      console.log("noping out of: \"" + threads[i].getFirstMessageSubject() + "\"");
      continue;
    }
    var tagName = "EvilEmail/" + mapAddrToTag[toHeader];
    var tag = GmailApp.getUserLabelByName(tagName);
    if (tag == null) {
      tag = GmailApp.createLabel(tagName);
    }
    tag.addToThread(threads[i]);
    
    tbpLabel.removeFromThread(threads[i]);
  }
}
