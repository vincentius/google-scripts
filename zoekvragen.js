function maakZoekvragenSheetOpDonderdag() { 
  Logger.log(SpreadsheetApp.getActiveSheet());
  
  var googleDoc = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1KYLnckVKGxAcwqumgGjj1FRSy5fDVIu6brILqwS6nSI/edit");
  
  var configSheet = googleDoc.getSheetByName("_CONFIG");
  
  // Doe niets als sheet niet actief is
  if ("X" != configSheet.getRange("B1").getValue().toUpperCase()) {
    Logger.log("Sheet niet actief")
    return;
  }
  
  var templateSheet = googleDoc.getSheetByName("_TEMPLATE");
  var templateRange = templateSheet.getRange("A1:Z50");
  
  for each (var sheet in googleDoc.getSheets()) {
    sheet.protect();
  }
  
  var d = new Date();
  var nextThursday = new Date();
  nextThursday.setDate(d.getDate() + ((7-d.getDay())%7+4) % 7);
  
  var nextThursdayString = Utilities.formatDate(nextThursday, "GMT+1", "dd/MM/yyyy");
  
  var newSheet = googleDoc.insertSheet(nextThursdayString);
  var newSheetRange = newSheet.getRange("A1:Z50");
  newSheet.activate();
  googleDoc.moveActiveSheet(0);
  
  templateRange.copyTo(newSheetRange);
  templateRange.copyTo(newSheetRange, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
  
  var shareLink = (
    "https://docs.google.com/spreadsheets/d/1KYLnckVKGxAcwqumgGjj1FRSy5fDVIu6brILqwS6nSI/edit#gid=" +
    newSheet.getSheetId() +
    " (Dit was een automatisch bericht)");
  
  Logger.log("Sharing Link: " + shareLink);
  
  var emails = configSheet.getRange("D1:D50").getValues().filter(function(x){ return x != ""; });
  
  for each (var email in emails){
    MailApp.sendEmail(email, "[BNI LVA] Deellink Google Sheet " + nextThursdayString, shareLink);
  } 
}
