/*
function onOpen() {
  var app = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  app.createMenu("Ranges")
  .addItem("Set ranges", "setRanges")
  .addItem("Delete ranges", "deleteRanges")
  .addToUi();
}

function setRanges() {
  var app = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var text = app.prompt("Give string", "Like A1:A10, Sheet2!A4:B5, This is a test!A:A", app.ButtonSet.OK_CANCEL); 
  if(text.getSelectedButton() == app.Button.OK) {
    var ranges = text.getResponseText().split(",");  
    
    for(var i = 0, iLen = ranges.length; i < iLen; i++) {
      var value, range, index, sname, split, sh;
      value = ranges[i].trim(), split = value.split("!");
      
      switch(split.length) {
        case 1:
          range = value;
          sh = ss.getActiveSheet();
          break;
        case 2:
          range = value;
          sh = ss.getSheetByName(String(split[0].trim()));
          break;
        default:
          index = value.lastIndexOf("!");
          range = "'" + value.slice(0, index) + "'!" + split[split.length - 1].trim();
          sh = ss.getSheetByName(String(value.slice(0, index)));
          break; 
      }
      
      try {
        sh.getRange(String(range))
        .protect()
        .setDescription("Range" + i);
      } catch(e) {
        throw "Range: " + (i) + " doesn't exist";
      }
    }
  }
}*/

