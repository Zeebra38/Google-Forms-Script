function setUpTrigger() {
  var spreadsheetsId = '1XT6aHEvD9AZvar8ypWYEFtibHGk-s6Ojx_Nmv04iK-A'.split(" ");
  spreadsheetsId.forEach(function (element) {
    if (ScriptApp.getUserTriggers(SpreadsheetApp.openById(element))) {
      var triggers = ScriptApp.getProjectTriggers();
      for (var i = 0; i < triggers.length; i++) {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    // ScriptApp.newTrigger('onEdit').forSpreadsheet(element).onEdit().create();
    ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(element).onFormSubmit().create();
    // ScriptApp.newTrigger('onChange').forSpreadsheet(element).onChange().create();
  })

}


function onFormSubmit(e) {
  var response = e;
  var ss = SpreadsheetApp.getActive();
  var name = ss.getName();
  var docFolderName = name.replace('(Ответы)', 'Документы');
  var curFile = DriveApp.getFileById(ss.getId());
  var folderId = curFile.getParents().next().getId();
  var dir = DriveApp.getFolderById(folderId);
  var docFolder = dir.getFoldersByName(docFolderName).next();
  var previousDocs = docFolder.searchFiles(`title contains "Записка №${namedValues['row']}"`);
  var firstTime = !previousDocs.hasNext();
  console.log(response);
  var namedValues;
  if (firstTime) {
    namedValues = response.namedValues;
  }
  else {

  }
  var adds = applicationSplit(getApplication(ss, response['range']['rowEnd']));
  namedValues['files Id'] = adds['filesId'];
  namedValues['Название приложения'] = adds['names'].join(" ");
  namedValues['row'] = response['range']['rowEnd'];
  namedValues['Номер'] = namedValues['row'];
  appendEditorUrl(ss, namedValues['row']);
  while (previousDocs.hasNext()) {
    Drive.Files.remove(previousDocs.next().getId());
  }
  var docID = createDoc(response.namedValues);
  var formName = FormApp.openByUrl(SpreadsheetApp.getActive().getFormUrl()).getTitle();
  var doc = DriveApp.getFileById(docID);
  var docs = [doc];
  for (let addId of adds['filesId']) {
    docs.push(DriveApp.getFileById(addId));
  }
  console.log(namedValues);
  responseToRespondent(getRespondentEmail(ss, namedValues['row']), "Форма принята к рассмотрению", formName, docs);
  sendOnApprove(SpreadsheetApp.getActive(), namedValues['row'], firstTime);
}

function applicationSplit(application) {
  if (application != "") {
    var adds = [];
    var filesId = [];
    var namedValues = {}
    application = application.replace(",", '').split(" ");
    for (let add of application) {
      filesId.push(add.split("id=")[1]);
      adds.push(DriveApp.getFileById(add.split("id=")[1]).getName());
    }
    namedValues['filesId'] = filesId;
    namedValues['names'] = adds;
    return namedValues;
  }
  return { 'filesId': [], 'names': [] }
}

function onEdit(e) {
  var response = e;
  console.log(response);
  if (response.value)
    console.log(response.value);
}

function onChange(e) {
  var response = e;
  console.log(response);
  if (response.value)
    console.log(response.value);
}

function getApplication(ss, findRow) {
  var sheet = ss.getActiveSheet();
  var lastCol = ss.getLastColumn();
  var range = sheet.getRange(1, 1, 1, lastCol);
  var values = range.getValues();
  for (var row in values) {
    for (var col in values[row]) {
      if (values[row][col] == "Приложение") {
        return sheet.getRange(findRow, parseInt(col) + 1).getValue();
      }
    }
  }
}
function createDoc(namedValues) {
  var ss = SpreadsheetApp.getActive();
  var name = ss.getName();
  var templateName = name.replace('(Ответы)', '(Шаблон)');
  var docFolderName = name.replace('(Ответы)', 'Документы');
  var template;
  var id = ss.getId();
  var curFile = DriveApp.getFileById(id);
  var folderId = curFile.getParents().next().getId();
  var dir = DriveApp.getFolderById(folderId);
  if (DriveApp.getFilesByName(templateName).hasNext()) {
    var file = DriveApp.getFilesByName(templateName).next();
    template = DocumentApp.openById(file.getId());
  }
  else {
    template = dir.createFile(templateName, "Измените этот шаблон для использования");
  }
  var docFolder;
  if (dir.getFoldersByName(docFolderName).hasNext()) {
    docFolder = dir.getFoldersByName(docFolderName).next();
  }
  else {
    docFolder = dir.createFolder(docFolderName);
  }
  var today = new Date().toLocaleString('ru', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: 'numeric',
    minute: 'numeric',
    second: 'numeric'
  });
  var todayData = today.split(" ");
  namedValues['День'] = todayData[0];
  namedValues['Месяц'] = todayData[1];
  namedValues['Год'] = todayData[2];
  var name = `${namedValues['Год']}-${new Date().getMonth() + 1}-${namedValues['День']} Записка №${namedValues['row']}.docx`;
  var docFile = DriveApp.getFileById(template.getId()).makeCopy(name, docFolder);
  var doc = DocumentApp.openById(docFile.getId());
  replaceValues(doc, namedValues);
  doc.saveAndClose();
  Utilities.sleep(500);
  doc = DocumentApp.openById(docFile.getId());
  saveAsPDF(doc, docFolder);
  return docFile.getId();
}

function getResponseValues(ss, row) {
  ss = SpreadsheetApp.openById("1XT6aHEvD9AZvar8ypWYEFtibHGk-s6Ojx_Nmv04iK-A");
  row = 16;
  var formURL = ss.getFormUrl();
  var form = FormApp.openByUrl(formURL);
  var sheet = ss.getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var columnIndex = headers[0].indexOf('Edit URL') + 1;
  var values = sheet.getRange(row, 1, 1, columnIndex - 1).getValues()[0];
  console.log(values);
  var formSubmitted = form.getResponses(values[0])[0];
  var itemResponses = formSubmitted.getGradableItemResponses();
  var namedValues = {}
  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];
    namedValues[itemResponse.getItem().getTitle()] = itemResponse.getResponse();
  }
  console.log(namedValues);
}
function replaceValues(doc, namedValues) {
  var body = doc.getBody();
  Object.entries(namedValues).forEach(function ([key, value]) {
    var pattern = "<<" + key + ">>";
    body.replaceText(pattern, value);
  });
}

function saveAsPDF(doc, folder) {
  var docPDF = doc.getAs('application/pdf');
  docPDF.setName(doc.getName().replace('.docx', '.pdf'));
  folder.createFile(docPDF);
}
