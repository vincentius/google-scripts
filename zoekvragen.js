function maakZoekvragenSheetOpDonderdag() { 
  Logger.log(SpreadsheetApp.getActiveSheet());
  
  var googleDoc = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/<knip>/edit");
  
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
    "https://docs.google.com/spreadsheets/d/<knip>/edit#gid=" +
    newSheet.getSheetId() +
    " (Dit was een automatisch bericht)");
    
  var emails = [
  // emails go here
  ];
  
  for each (var email in emails){
    MailApp.sendEmail(email, "[BNI LVA] Deellink Google Sheet " + nextThursdayString, shareLink);
  } 
}
