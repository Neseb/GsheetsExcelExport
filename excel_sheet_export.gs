
// Copie les lignes non commandées dans une feuille temporaire et met a jour la feuille originelle

function emailNotOrdered(){

  var receipients = "marianne.burbage@curie.fr, sandrine.heurtebise@curie.fr, nina.burgdorf@curie.fr";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = SpreadsheetApp.setActiveSheet(ss.getSheets()[0])
  var col = 15 ; // choose the column holding the status. 0-indexing
  var range = sh.getDataRange();
  var values = range.getValues();// data is a 2D array, index0 = col A
  var target= new Array();// this is a new array to collect data
  var date = new Date();
  var day = (date.getDate() < 9 ? '0': '') + date.getDate();
  var monthIndex = (date.getMonth() < 9 ? '0': '') + (date.getMonth()+1);
  var year = date.getFullYear();

  var corrected = day + '/' + monthIndex + '/' + (year-2000);

  // Copy the headers
  target.push(values[0].slice(1,col));

  // Skip the first line
  for(n=1;n<range.getHeight();++n){
    var row = values[n];
    if (row[col]!='Commandé'){
      // skip first column
      target.push(row.slice(1,col));
      values[n][col] = 'Commandé';
      values[n][col+1] = corrected;
       }
   }

   if(target.length>1){// if there is something to copy

    var exportName = "Commandes " + corrected;

     var ssTmp = SpreadsheetApp.create(exportName);
     SpreadsheetApp.setActiveSpreadsheet(ssTmp);
     var shTmp = SpreadsheetApp.setActiveSheet(ssTmp.getSheets()[0]);
     shTmp.clear();
     shTmp.getRange(1,1,target.length,target[0].length).setValues(target);

     SpreadsheetApp.flush();

     var  idTmp = ssTmp.getId();

     var spreadsheet   = SpreadsheetApp.getActiveSpreadsheet();
     var spreadsheetId = spreadsheet.getId()
     var file          = Drive.Files.get(spreadsheetId);
     var url           = file.exportLinks[MimeType.MICROSOFT_EXCEL];
     var token         = ScriptApp.getOAuthToken();
     var response      = UrlFetchApp.fetch(url, {
       headers: {
         'Authorization': 'Bearer ' +  token
       }
     });

     var fileName = (spreadsheet.getName()) + '.xlsx';
     var blobs   = [response.getBlob().setName(fileName)];
     var subject = exportName;
     var emailbody = exportName;

     MailApp.sendEmail(receipients, subject, emailbody, {attachments: blobs});

     SpreadsheetApp.setActiveSpreadsheet(ss);
     sh.getRange(1,1,values.length,values[0].length).setValues(values);

     Drive.Files.remove(idTmp);

   }  else {
     var noOrders = "Pas de commandes aujourd'hui (" + corrected + ")";
     MailApp.sendEmail(receipients, noOrders, "Aucune commande n'a été ajoutée depuis la dernière fois.");

   }
}

