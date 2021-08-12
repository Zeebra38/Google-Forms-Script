/**
 * Проверяет наличие триггеров на Таблицах по их Id, разделенных "/n". Если они есть, то удаляет, а потом устанавливает триггер "On form submit" на функицю "onFormSubmit"
 */
function setUpTrigger(spreadsheetsId) {
  spreadsheetsId = spreadsheetsId.split("\n");
  spreadsheetsId.forEach(function (element) {
    if (element != undefined && element.length > 5) {
      if (ScriptApp.getUserTriggers(SpreadsheetApp.openById(element))) {
        var triggers = ScriptApp.getProjectTriggers();
        for (var i = 0; i < triggers.length; i++) {
          ScriptApp.deleteTrigger(triggers[i]);
        }
      }
      ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(element).onFormSubmit().create();
    }
  })

}
/**
 * Устанавливает триггеры на все Таблицы по их Id, разделенные '\n'
 */
function mainSetup(ssId = "") {
  setUpTrigger(ssId)
}

/**
 * Основная фукнция программы. Срабатывает при отправке формы
 * @param {event} e Событие, которое передает Google при вызове функции. Внутри namedValues - объект, который создержит пары ключ-значение ответов формы.
 */
function onFormSubmit(e) {
  var response = e;
  var ss = SpreadsheetApp.getActive();
  var name = ss.getName();
  var docFolderName = name.replace('(Ответы)', 'Документы');
  var curFile = DriveApp.getFileById(ss.getId());
  var folderId = curFile.getParents().next().getId();
  var dir = DriveApp.getFolderById(folderId);
  var docFolder = dir.getFoldersByName(docFolderName).next();
  var previousDocs = docFolder.searchFiles(`title contains "Записка №${response['range']['rowEnd']}"`);
  var firstTime = !previousDocs.hasNext();
  var namedValues;
  if (firstTime) {
    namedValues = response.namedValues;
  }
  else {
    namedValues = getResponseValues(ss, response['range']['rowEnd']);
  }
  var adds = applicationSplit(getApplication(ss, response['range']['rowEnd']));
  namedValues['files Id'] = adds['filesId'];
  namedValues['Название приложения'] = adds['names'].join(", ");
  namedValues['row'] = response['range']['rowEnd'];
  namedValues['Номер'] = namedValues['row'];
  appendEditorUrl(ss, namedValues['row']);
  while (previousDocs.hasNext()) {
    Drive.Files.remove(previousDocs.next().getId());
  }
  var namedValuesMap = new Map(Object.entries(namedValues));
  for (let key of namedValuesMap.keys()) {
    if (key.indexOf("вложение при необходимости") != -1) {
      if (namedValuesMap.get(key.split(" ").slice(0, -3).join(" ").trim()) == "" && namedValuesMap.get(key) != "") {
        var bufArray = [];
        var curArray = namedValuesMap.get(key);
        for (var i = 0; i < curArray.length; i++) {
          if (!firstTime) {
            bufArray.push(`https://drive.google.com/open?id=${curArray[i]}`);
          }
          else {
            bufArray.push(curArray[i]);
          }
        }
        namedValuesMap.set(key.split(" ").slice(0, -3).join(" ").trim(), bufArray.join(", "));
      }
    }
  }
  namedValues = Object.fromEntries(namedValuesMap);
  var docID = createDoc(namedValues);
  var formName = FormApp.openByUrl(ss.getFormUrl()).getTitle();
  var doc = DriveApp.getFileById(docID);
  var docs = [doc];
  for (let addId of adds['filesId']) {
    docs.push(DriveApp.getFileById(addId));
  }
  if (firstTime) {
    responseToRespondent(getRespondentEmail(ss, namedValues['row']), "направлена на согласование", formName, docs, "", "", ss, namedValues['row']);
  }
  if (checkFinalStage(ss, namedValues['row'])) {
    sendToDestination(ss, namedValues['row']);
  }
  else {
    sendOnApprove(ss, namedValues['row'], firstTime);
  }
}

/**
 * Функция разделяет Приложения на 2 массива, которые записаны в Объект. filesId - Id файлов приложений, names - имена файлов приложений
 * @param {string} application Строка, содержащая ссылки на приложения, записанные через ", "
 */
function applicationSplit(application) {
  if (application != undefined && application.trim() != "") {
    var adds = [];
    var filesId = [];
    var namedValues = {}
    application = application.replaceAll(",", '').split(" ");
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

/**
 * Находит данные ячейки по findRow и первому столбцу, содержащему слово "приложение"\
 * @param {Spreadsheet} ss Гугл Таблица
 * @param {number} findRow Номер строки
 */
function getApplication(ss, findRow) {
  var sheet = ss.getActiveSheet();
  var lastCol = ss.getLastColumn();
  var range = sheet.getRange(1, 1, 1, lastCol);
  var values = range.getValues();
  for (var row in values) {
    for (var col in values[row]) {
      if (values[row][col].toLowerCase().indexOf("приложение") != -1) {
        return sheet.getRange(findRow, parseInt(col) + 1).getValue();
      }
    }
  }
}

/**
 * Создает документ из шаблона "%Название формы% (Шаблон)"
 * @param {object} namedValues Словарь ключ-значение, содержащий данные
 */
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
  Utilities.sleep(500);
  try {
    doc.saveAndClose();
  }
  catch (exception) {
    console.log(exception);
    Drive.Files.remove(doc.getId());
    createDoc(namedValues);
  }
  Utilities.sleep(500);
  doc = DocumentApp.openById(docFile.getId());
  saveAsPDF(doc, docFolder);
  return docFile.getId();
}
/**
 * Если форма была отправлена раньеш, то при редактировании формы в namedValues будут только измененные пункты. Эта фукнция находит данные прошлых ответов формы по timestamp (1 столбец в таблице)
 * @param {Spreadsheet} ss Гугл Таблица
 * @param {number} row Номер строки
 */
function getResponseValues(ss, row) {
  var formURL = ss.getFormUrl();
  var form = FormApp.openByUrl(formURL);
  var sheet = ss.getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var columnIndex = headers[0].indexOf('Edit URL') + 1;
  var values = sheet.getRange(row, 1, 1, columnIndex - 1).getValues()[0];
  var formSubmitted = form.getResponses(values[0])[0];
  var itemResponses = formSubmitted.getGradableItemResponses();
  var namedValues = {}
  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];
    namedValues[itemResponse.getItem().getTitle()] = itemResponse.getResponse();
  }
  return namedValues;
}

/**
 * Функция, используемая для замены данных в документе, находящихся в виде "<<%Название пункта в форме%>>"
 * @param {Document} doc открытый Гугл Документ
 * @param {object} namedValues Словарь ключ-значение, содержащий данные
 */
function replaceValues(doc, namedValues) {
  var body = doc.getBody();
  Object.entries(namedValues).forEach(function ([key, value]) {
    var pattern = "<<" + key + ">>";
    body.replaceText(pattern, value);
  });
}

/**
 * Функция, используемая для сохранения Документа в виде .pdf
 * @param {Document} doc открытый Гугл Документ
 * @param {Folder} folder Папка, куда необходимо сохранить итоговый .pdf
 */
function saveAsPDF(doc, folder) {
  var docPDF = doc.getAs('application/pdf');
  docPDF.setName(doc.getName().replace('.docx', '.pdf'));
  folder.createFile(docPDF);
}
