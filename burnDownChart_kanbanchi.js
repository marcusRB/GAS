function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('UX - Custom Menu')
      .addItem('UpdateNow!', 'driveActivityReport')
      .addToUi();
}


function driveActivityReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  //var sheet = ss.getActiveSheet();
  var sheetName = "sheetGenerated";
  //var sheet = ss.getSheetByName(sheetName).getRange(2,2);
  var sheet = ss.getSheetByName(sheetName).hideSheet();
  
  
  // Recupera il timezone dello Spreadsheet"
  var timezone = ss.getSpreadsheetTimeZone();
  
  var today     = new Date();
  var oneDayAgo = new Date(today.getTime() - 1 * 1 * 60 * 60 * 1000);
  var startTime = oneDayAgo.toISOString();
  
  
  // Trova i file modificati nelle ultime 24 ore
  var search = 'modifiedDate > "' + startTime + '" and title contains "UX TAREAS SPRINT " and mimeType contains "spreadsheet"';
  var files  = DriveApp.searchFiles(search);
  
  //var sharedFile = files.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
   
  // Effettua un ciclo su tutti i file recuperati dal criterio di ricerca utilizzato
  while( files.hasNext() ) {  
    var file = files.next();
    
    var fileName     = file.getName();
    var fileID       = file.getId();
    //var fileURL      = file.getUrl();
    var sheetID      = file.getId();
    var dateCreated  = Utilities.formatDate(file.getDateCreated(), timezone, "yyyy-MM-dd HH:mm");
    var dateModified = Utilities.formatDate(file.getLastUpdated(), timezone, "yyyy-MM-dd HH:mm");
    // Appende una riga allo Spreadsheet con le info del file
    //sheet.appendRow([fileName, fileURL, dateCreated, dateModified, sheetID]);
    //sheet.setValue(sheetID);
    sheet.getRange(2, 1, 1, 4).setValues([[fileName, sheetID, dateCreated, dateModified]]);
    
    //Abrimos la spreadsheet, seleccionamos la hoja
    var spreadsheet = SpreadsheetApp.openById(fileID);
    var sheet1 = spreadsheet.getSheets()[0];

    //Seleccionamos el rango  
    var range = sheet1.getRange("A1:M100");
    values = range.getValues();

    //Seleccionamos la hoja de destino, que es la activeSheet 

    var SSsheet = ss.getSheetByName('testing'); // ts = target sheet;

  //Seleccionamos el mismo rango y le asignamos los valores
  var ssRange = SSsheet.getRange("A1:M100")
  ssRange.setValues(values);
  SSsheet.setTabColor("ff0000");
  var column = SSsheet.getRange("G2:G");
  column.setNumberFormat("dd-MMM");
  SSsheet.hideSheet();
  
  ss.getSheetByName('dashboard').activate();
    
   }

}





