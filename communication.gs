class Approver {
  constructor(email, name, column) {
    this.name = name;
    this.email = email;
    this.column = column;
  }
}

function sendEmail(email, subj, message) {
  console.log(email);
  MailApp.sendEmail(
    {
      to: email,
      subject: subj,
      htmlBody: message
    }
  );
}

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


function responseToRespondent(email, subject, formName, files, comments = "", editUrl = "") {
  var templ;
  switch (subject) {
    case "Форма отклонена":
      templ = HtmlService.createTemplateFromFile('rejectedForm');
      templ.comments = comments;
      templ.editUrl = editUrl;
      break;
    case "Форма одобрена":
      templ = HtmlService.createTemplateFromFile('approvedForm');
      templ.comments = comments;
      break;
    case "Форма принята к рассмотрению":
      templ = HtmlService.createTemplateFromFile('starterNotification');
      break;
  }
  templ.formName = formName;
  var message = templ.evaluate().getContent();
  sendEmailWithAttach(email, subject, message, files);
}

function getListOfApprovers(ss) {
  // ss = SpreadsheetApp.openById("1XT6aHEvD9AZvar8ypWYEFtibHGk-s6Ojx_Nmv04iK-A");
  var sheet = ss.getActiveSheet();
  var lastCol = ss.getLastColumn();
  var range = sheet.getRange(1, 1, 1, lastCol);
  var values = range.getValues();
  var pattern = /\S+@\S+\/.*/;
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
 * @param {Boolean} firstTime
 */
function sendOnApprove(ss, row, firstTime) {
  // ss = SpreadsheetApp.openById("1XT6aHEvD9AZvar8ypWYEFtibHGk-s6Ojx_Nmv04iK-A");
  // row = 17;
  var sheet = ss.getActiveSheet();
  var name = ss.getName();
  console.log(firstTime);
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
  getListOfApprovers(ss).forEach(function (approver) {
    var cell = sheet.getRange(row, approver.column);
    if (firstTime) {
      cell.setValue("На обработке");
    }
    if (!sended) {
      if (cell.getValue() == "На обработке" || cell.getValue() == "Отклонено") {
        var templ = HtmlService.createTemplateFromFile('goToApprove');
        templ.row = row;
        templ.column = approver.column;
        templ.ssID = ss.getId();
        templ.scriptURL = loadSettings().scriptURL;
        var message = templ.evaluate().getContent();
        sendEmailWithAttach(approver.email, "Подтвердите форму", message, docs);
        sended = true;
      }
    }
  });
}

function readyCheck(ss, row, column) {
  var sheet = ss.getActiveSheet();
  var lastCol = ss.getLastColumn();
  var approvers = getListOfApprovers(ss);
  var start_responses = [];
  var notes = [];
  console.log(approvers);
  for (let approver of approvers) {
    start_responses.push(sheet.getRange(row, approver.column).getValue());
    notes.push(sheet.getRange(row, approver.column).getNote());
  }
  var responses = [...new Set(start_responses)];
  console.log(responses);
  var name = ss.getName();
  var docFolderName = name.replace('(Ответы)', 'Документы');
  var curFile = DriveApp.getFileById(ss.getId());
  var folderId = curFile.getParents().next().getId();
  var dir = DriveApp.getFolderById(folderId);
  var docFolder = dir.getFoldersByName(docFolderName).next();
  var doc = docFolder.searchFiles(`title contains "Записка №${row}"`).next();
  var docs = [doc]
  var addStr = getApplication(ss, row);
  var adds = applicationSplit(addStr);
  for (let addId of adds['filesId']) {
    docs.push(DriveApp.getFileById(addId));
  }
  var comment = notes.join(";                 ");
  var formName = FormApp.openByUrl(ss.getFormUrl()).getTitle();
  if (responses.indexOf("Отклонено") != -1) {
    responseToRespondent(getRespondentEmail(ss, row), "Форма отклонена", formName, docs, comment, appendEditorUrl(ss, row, true));
  }
  else if (responses.length == 1 && responses.indexOf("Подтверждено") != -1) {
    var formName = FormApp.openByUrl(ss.getFormUrl()).getTitle();
    responseToRespondent(getRespondentEmail(ss, row), "Форма одобрена", formName, docs, comment);
  }
  else if(responses.indexOf("Подтверждено") != -1)
  {
    sendOnApprove(ss, row, false);
  }
}

function getRespondentEmail(ss, number) {
  var sheet = ss.getActiveSheet();
  var lastCol = ss.getLastColumn();
  var range = sheet.getRange(1, 1, 1, lastCol);
  var values = range.getValues();
  var column = values[0].indexOf('Адрес электронной почты');
  return sheet.getRange(number, parseInt(column) + 1).getValue();
}

function appendEditorUrl(ss, row, justReturn = false) {
  // ss = SpreadsheetApp.openById("1XT6aHEvD9AZvar8ypWYEFtibHGk-s6Ojx_Nmv04iK-A");
  // row = 16;
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

function test() {
  var ss = SpreadsheetApp.openById("1XT6aHEvD9AZvar8ypWYEFtibHGk-s6Ojx_Nmv04iK-A");
  var sheet = ss.getActiveSheet();
  console.log(sheet.getLastRow());
}