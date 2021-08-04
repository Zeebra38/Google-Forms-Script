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
  console.log(response);
  if (response.namedValues) {
    console.log(response.namedValues);
  }
  var namedValues = response.namedValues;
  namedValues['row'] = response['range']['rowEnd'];
  namedValues['Номер'] = namedValues['row'];
  var docID = createDoc(response.namedValues);
  var formName = FormApp.openByUrl(SpreadsheetApp.getActive().getFormUrl()).getTitle();
  responseToRespondent(namedValues['Адрес электронной почты'][0], "Форма принята к рассмотрению", formName, DriveApp.getFileById(docID));
  sendOnApprove(SpreadsheetApp.getActive(), namedValues['row']);
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
  console.log(name);
  docFile = DriveApp.getFileById(template.getId()).makeCopy(name, docFolder);
  var doc = DocumentApp.openById(docFile.getId());
  replaceValues(doc, namedValues);
  doc.saveAndClose();
  var doc = DocumentApp.openById(docFile.getId());
  saveAsPDF(doc, docFolder);
  Utilities.sleep(2000);
  return docFile.getId();
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
