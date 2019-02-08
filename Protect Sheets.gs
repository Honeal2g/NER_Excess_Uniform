/*function onOpen() {
var app = SpreadsheetApp.getUi();
app.createMenu("Permissions")
.addItem("Set POC Rights", "SetPOCPermissions")
.addToUi();
.addItem("Set Admin Rights", "SetADMINPermissions") 
//.addItem("Clear Protections", "RemoveProtection")

}*/
function SetADMINPermissions() {
  var app = SpreadsheetApp.getUi();
  
  var input = app.prompt("Select DET to Assign Permissions", "For example, DET 9, DET 560", app.ButtonSet.OK_CANCEL);
  
  var DetName = input.getResponseText();
  
  var Sheets = ['DET 9','DET 115','DET 128','DET 130','DET 215','DET 220','DET 225','DET 330','DET 340','DET 355','DET 365','DET 370','DET 380','DET 390','DET 475',
                'DET 485','DET 490','DET 520','DET 535','DET 536','DET 538','DET 550','DET 560','DET 590','DET 620','DET 630','DET 640','DET 643','DET 650','DET 665',
                'DET 720','DET 730','DET 750','DET 752','DET 867','DET 915'];  
  
  var AdminLock = ['A9:B169','F9:G169','K9:L98','A5:G7','A1:I1','A3:I3','K1:N1','H9:N9','H24:I24','H40:I40','H43:I43','H49:I49','H65:I65','H81:I81','H96:I96','H112:I112',
                   'H128:I128','H139:I139','H148:I148','H154:I154','H166:I166','H168:I168','C154:D154','C139:D139','C124:D124','C110:D110','C94:D94','C82:D82','C70:D70',
                   'C58:D58','C46:D46','C35:D35','C24:D24','C9:D9','M25:N25','M41:N41','M48:N48','M57:N57','M66:N66','M82:N82','M98:N98'];
  
  var Master_Roster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Roster').getRange('B2:B37').getValues(); //Email Array
  
  if(input.getSelectedButton() == app.Button.OK) { //if OK was selected then proceed
    for(var i = 0; i < Sheets.length; i++){
      if(DetName == String(Sheets[i])){ 
        for(var x = 0; x < 41; x++){ //Loops for Number of Ranges to Protect
          var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Sheets[i]); 
          var range = ss.getRange(Sheets[i]+'!'+AdminLock[x]);
          var protection = range.protect().setDescription(Sheets[i]+' Admin Permissions '+x+'/40');
          protection.addEditor(Master_Roster[i]);//This is a placeholder as this line is superceded by the line of code below.    
          protection.removeEditors(protection.getEditors());//Removes all except Team Member with FULL rights (Owner)
          if (protection.canDomainEdit()) {
            protection.setDomainEdit(false);
          }
        }
      }
    } 
  }  
}
function SetPOCPermissions() {
  var app = SpreadsheetApp.getUi(); 
  
  var input = app.prompt("Select DET to Assign Permissions", "For example, DET 9, DET 560", app.ButtonSet.OK_CANCEL);
  
  var DetName = input.getResponseText();
  
  var Sheets = ['DET 9','DET 115','DET 128','DET 130','DET 215','DET 220','DET 225','DET 330','DET 340','DET 355','DET 365','DET 370','DET 380','DET 390','DET 475',
                'DET 485','DET 490','DET 520','DET 535','DET 536','DET 538','DET 550','DET 560','DET 590','DET 620','DET 630','DET 640','DET 643','DET 650','DET 665',
                'DET 720','DET 730','DET 750','DET 752','DET 867','DET 915'];  
  
  var UsrLock = ['C10:D23','C25:D34','C36:D45','C47:D57','C59:D69','C71:D81','C83:D93','C95:D109','C111:D123','C125:D138','C140:D153','C155:D169','H10:I23','H25:I39',
                 'H41:I42','H44:I48','H50:I64','H66:I80','H82:I95','H97:I111','H113:I127','H129:I138','H140:I147','H149:I153','H155:I165','H167:I167','H169:I169',
                 'M10:N24','M26:N40','M42:N47','M49:N56','M58:N65','M67:N81','M83:N97','K99:N169','K2:N7','A2:I2','H5:I7'];
  
  var Master_Roster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Roster').getRange('B2:B37').getValues(); //Email Array
  
  if(input.getSelectedButton() == app.Button.OK) { //if OK was selected then proceed
    for(var n = 0; n < Sheets.length; n++){
      if(DetName == String(Sheets[n])){ 
        for(var i = 0; i < 1; i++){//Loops for Number of Dets
          for(var x = 0; x < 38; x++){ //Loops for Number of Ranges to Protect
            var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Sheets[n]);
            var range = ss.getRange(Sheets[n]+'!'+UsrLock[x]);
            var protection = range.protect().setDescription(Sheets[n]+' POC Permissions '+x+'/37'); //Discription Nomenclature 
            protection.addEditor(Master_Roster[n]);//User to give rights      
          }
        }
      }
    }
  }
}

function RemoveAllProtections() {
  // Remove all range protections in the spreadsheet that the user has permission to edit.
  var ss = SpreadsheetApp.getActive();
  var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    if (protection.canEdit()) {
      protection.remove();
    }
  }
}
function RemoveSheetProtection(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0]; // assuming you want the first sheet
  var protections = sheet.getProtections();
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() == 'Protect column A') {
      protection[i].remove();
    }
  }
}  
function resetPageLayout() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Sheet1');
  ss.toast('Now processing your sheet','Wait a few seconds',5);
  if(sh.getMaxRows()-sh.getLastRow()>0){sh.deleteRows(sh.getLastRow()+1, sh.getMaxRows()-sh.getLastRow())};
  if(sh.getMaxColumns()-sh.getLastColumn()>0){sh.deleteColumns(sh.getLastColumn()+1, sh.getMaxColumns()-sh.getLastColumn())};
  var sheets = ss.getSheets();
  for(var n=0;n<sheets.length;n++){
    if(sheets[n].getName()!='Sheet1'){
      try{
        ss.deleteSheet(sheets[n])}catch(err){
          Browser.msgBox('Can\'t delete Sheet named "'+sheets[n].getName()+'" ('+err+')');
        }
    }
  }
  SpreadsheetApp.flush();
}
//Deletes All the Empty Row to Clean up the sheets appearance
function removeEmptyRows() {
  var ss = SpreadsheetApp.getActive();
  var allsheets = ss.getSheets();
  for (var s in allsheets){
    var sheet=allsheets[s]
    var maxRows = sheet.getMaxRows(); 
    var lastRow = sheet.getLastRow();
    if (maxRows-lastRow != 0){
      sheet.deleteRows(lastRow+1, maxRows-lastRow);
    }
  }
}//end Remove empty row