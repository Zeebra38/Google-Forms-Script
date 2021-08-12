/**
 * Класс для Одобрителя
 */
class Approver {
  constructor(email, name, column) {
    this.name = name;
    this.email = email;
    this.column = column;
  }
}
/**
 * Функция для отправки email с прикрепленными файлами
 * @param {string} email Электронный адрес получателя
 * @param {string} subj Тема письма
 * @param {htmlOutput} message Html содержание сообщения
 * @param {[File]} files Массив прикрепленных файлов
 */
function sendEmailWithAttach(email, subj, message, files) {
  MailApp.sendEmail(
    {
      to: email,
      subject: subj,
      htmlBody: message,
      attachments: files
    }
  );
}

/**
 * Отвечает респонденту
 * @param {string} email Электронный адрес получателя
 * @param {string} subject Тема письма
 * @param {string} formName Заголовок формы, на которую он отвечал
 * @param {[File]} files Массив прикрепленных файлов
 * @params {string} comments Комментарии "одобрителей"
 * @params {string} editUrl Ссылка на редактирование
 * @param {Spreadsheet} ss Гугл Таблица
 * @param {number} row Номер строки
 * @param {number} col Номер колонки; от кого пришел отказ
 */
function responseToRespondent(email, subject, formName, files, comments = "", editUrl = "", ss, row, col=0) {
  var templ;
  switch (subject) {
    case "отклонена":
      templ = HtmlService.createTemplateFromFile('rejectedForm');
      templ.whoRejected = ss.getActiveSheet().getRange(1, col).getValue().split('/')[0];
      templ.comments = comments;
      templ.editUrl = editUrl;
      break;
    case "одобрена":
      templ = HtmlService.createTemplateFromFile('approvedForm');
      templ.approvers = getApproversToEmail(ss);
      templ.comments = comments;
      break;
    case "направлена на согласование":
      templ = HtmlService.createTemplateFromFile('starterNotification');
      templ.approvers = getApproversToEmail(ss);
      break;
  }
  templ.formName = formName;
  var message = templ.evaluate().getContent();
  sendEmailWithAttach(email, getSubject(ss ,row) + subject, message, files);
}

/**
 * Возвращает список "одобртелей"
 * @param {Spreadsheet} ss Гугл Таблица
 */
function getListOfApprovers(ss) {
  var sheet = ss.getActiveSheet();
  var lastCol = ss.getLastColumn();
  var range = sheet.getRange(1, 1, 1, lastCol);
  var values = range.getValues();
  var pattern = /\S+@\S+\/\S.*/;
  var approvers = [];
  for (var row in values) {
    for (var col in values[row]) {
      if (values[row][col].match(pattern)) {
        var splitList = values[row][col].split("/");
        approvers.push(new Approver(splitList[0], splitList[1], parseInt(col) + 1));
      }
    }
  }
  return approvers;
}

/**
 * Заполняет стобцы "одобрителей" фразой "На обработке", а также отправляет письмо следующему по порядку отправителю
 * @param {Spreadsheet} ss Гугл Таблица
 * @param {number} row Номер строки, от которой пришел запрос с изменением содержимого
 * @param {boolean} firstTime Это первая отправка формы или ответ уже был изменен?
 */
function sendOnApprove(ss, row, firstTime) {
  var sheet = ss.getActiveSheet();
  var name = ss.getName();
  var docFolderName = name.replace('(Ответы)', 'Документы');
  var curFile = DriveApp.getFileById(ss.getId());
  var folderId = curFile.getParents().next().getId();
  var dir = DriveApp.getFolderById(folderId);
  var docFolder = dir.getFoldersByName(docFolderName).next();
  var doc = docFolder.searchFiles(`title contains "Записка №${row}"`).next();
  var docs = [doc];
  var addStr = getApplication(ss, row);
  var adds = applicationSplit(addStr);
  for (let addId of adds['filesId']) {
    docs.push(DriveApp.getFileById(addId));
  }
  var sended = false;
  var approvers = getListOfApprovers(ss);
  for(var i = 0; i < approvers.length - 1; i++)
  {
    var approver = approvers[i];
    var cell = sheet.getRange(row, approver.column);
    if (firstTime) {
      cell.setValue("На обработке");
    }
    if (!sended) {
      if (cell.getValue() == "На обработке" || cell.getValue() == "Отклонено") {
        var templ = HtmlService.createTemplateFromFile('goToApprove');
        templ.row = row;
        templ.formName = FormApp.openByUrl(ss.getFormUrl()).getTitle();
        templ.column = approver.column;
        templ.ssID = ss.getId();
        templ.scriptURL = loadSettings().scriptURL;
        var message = templ.evaluate().getContent();
        sendEmailWithAttach(approver.email, getSubject(ss, row) +  "необходимо рассмотреть", message, docs);
        sended = true;
      }
    }
  }
}

/**
 * Отправляет готовый одобренный весми документ адресату
 * @param {Spreadsheet} ss Гугл Таблица
 * @param {number} row Номер строки, от которой пришел запрос с изменением содержимого
 */
function sendToDestination(ss, row)
{
  var destinationApprover = getListOfApprovers(ss).pop();
  ss.getActiveSheet().getRange(row, destinationApprover.column).setValue("Отправлено");
  var name = ss.getName();
  var docFolderName = name.replace('(Ответы)', 'Документы');
  var curFile = DriveApp.getFileById(ss.getId());
  var folderId = curFile.getParents().next().getId();
  var dir = DriveApp.getFolderById(folderId);
  var docFolder = dir.getFoldersByName(docFolderName).next();
  var doc = docFolder.searchFiles(`title contains "Записка №${row}"`).next();
  var docs = [doc];
  var addStr = getApplication(ss, row);
  var adds = applicationSplit(addStr);
  for (let addId of adds['filesId']) {
    docs.push(DriveApp.getFileById(addId));
  }
  var templ = HtmlService.createTemplateFromFile("Result");
  var formName = FormApp.openByUrl(ss.getFormUrl()).getTitle();
  templ.formName = formName;
  templ.comments = createComment(ss, row);
  templ.editUrl = appendEditorUrl(ss, row, true);
  var message = templ.evaluate().getContent();
  sendEmailWithAttach(destinationApprover.email, getSubject(ss, row) + `по форме "${formName}"`,  message, docs);
}

/**
 * Проверка ответов "одобрителей". В зависимости от ответов передает управление дальше
 * @param {Spreadsheet} ss Гугл Таблица
 * @param {number} row Номер строки, от которой пришел запрос с изменением содержимого
 * @param {number} column Номер колонки, от которой пришел запрос с изменением содержимого
 */
function readyCheck(ss, row, column) {
  var sheet = ss.getActiveSheet();
  var lastCol = ss.getLastColumn();
  var approvers = getListOfApprovers(ss);
  var start_responses = [];
  for(var i = 0; i < approvers.length - 1; i ++)
  {
    start_responses.push(sheet.getRange(row, approvers[i].column).getValue()); 
  }
  var responses = [...new Set(start_responses)];
  var name = ss.getName();
  var docFolderName = name.replace('(Ответы)', 'Документы');
  var curFile = DriveApp.getFileById(ss.getId());
  var folderId = curFile.getParents().next().getId();
  var dir = DriveApp.getFolderById(folderId);
  var docFolder = dir.getFoldersByName(docFolderName).next();
  var doc = docFolder.searchFiles(`title contains "Записка №${row}"`).next();
  var docs = [doc];
  var addStr = getApplication(ss, row);
  var adds = applicationSplit(addStr);
  for (let addId of adds['filesId']) {
    docs.push(DriveApp.getFileById(addId));
  }
  var formName = FormApp.openByUrl(ss.getFormUrl()).getTitle();
  if (responses.indexOf("Отклонено") != -1) {
    responseToRespondent(getRespondentEmail(ss, row), "отклонена", formName, docs, createComment(ss, row), appendEditorUrl(ss, row, true), ss, row, column);
  }
  else if (responses.length == 1 && responses.indexOf("Подтверждено") != -1) {
    var formName = FormApp.openByUrl(ss.getFormUrl()).getTitle();
    sendToDestination(ss, row);
    responseToRespondent(getRespondentEmail(ss, row), "одобрена", formName, docs, createComment(ss, row), "", ss, row);
  }
  else if(responses.indexOf("Подтверждено") != -1)
  {
    sendOnApprove(ss, row, false);
  }
}
/**
 * Возращает email исполнителя. Он должен быть написан в столбце с названием "e-mail исполнителя"
 * @param {Spreadsheet} ss Гугл Таблица
 * @param {number} number Номер строки, от которой пришел запрос с изменением содержимого
 */
function getRespondentEmail(ss, number) {
  var sheet = ss.getActiveSheet();
  var lastCol = ss.getLastColumn();
  var range = sheet.getRange(1, 1, 1, lastCol);
  var values = range.getValues();
  var column = values[0].indexOf('e-mail исполнителя');
  return sheet.getRange(number, parseInt(column) + 1).getValue();
}

/**
 * Добавляет или просто возвращает (если justReturn = true) ссылку на редактирование таблицы
 * @param {Spreadsheet} ss Гугл Таблица
 * @param {number} row Номер строки, от которой пришел запрос с изменением содержимого
 * @params {bollean} justReturn Если true, то просто возвращает ссылку и ничего не делает с контентом таблицы
 */
function appendEditorUrl(ss, row, justReturn = false) {
  if (!justReturn) {
    var formURL = ss.getFormUrl();
    var form = FormApp.openByUrl(formURL);
    var sheet = ss.getActiveSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    var columnIndex = headers[0].indexOf('Edit URL') + 1;
    var values = sheet.getRange(row, 1, 1, columnIndex - 1).getValues()[0];
    var formSubmitted = form.getResponses(values[0]);
    var editResponseUrl = formSubmitted[0].getEditResponseUrl();
    sheet.getRange(row, columnIndex).setValue(editResponseUrl);
    return editResponseUrl;
  }
  else {
    var sheet = ss.getActiveSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    var columnIndex = headers[0].indexOf('Edit URL') + 1;
    return sheet.getRange(row, columnIndex).getValue();
  }
}