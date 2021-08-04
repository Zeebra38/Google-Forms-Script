class Approver
{
  constructor(email, name, column)
  {
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


function responseToRespondent(email, subject, formName, files, comments="") {
  var templ;
  switch (subject)
  {
    case "Форма отклонена":
    templ = HtmlService.createTemplateFromFile('rejectedForm');
    templ.comments = comments;
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
      if (values[row][col].match(pattern))
      {
        var splitList = values[row][col].split("/");
        approvers.push(new Approver(splitList[0], splitList[1], parseInt(col) + 1));
      }
    }
  }
  return approvers;
}

function sendOnApprove(ss, row) {
  // ss = SpreadsheetApp.openById("1XT6aHEvD9AZvar8ypWYEFtibHGk-s6Ojx_Nmv04iK-A");
  // row = 17;
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
  for (let addId of adds['filesId'])
  {
    docs.push(DriveApp.getFileById(addId));
  }
  getListOfApprovers(ss).forEach(function(approver) {
    sheet.getRange(row, approver.column).setValue("?");
    var templ = HtmlService.createTemplateFromFile('goToApprove');
    templ.row = row;
    templ.column = approver.column;
    templ.ssID = ss.getId();
    templ.scriptURL = loadSettings().scriptURL;
    var message = templ.evaluate().getContent();
    sendEmailWithAttach(approver.email, "Подтвердите форму", message, docs);
  });
}

function readyCheck(ss, row, column)
{
  var sheet = ss.getActiveSheet();
  var lastCol = ss.getLastColumn();
  var approvers = getListOfApprovers(ss);
  var start_responses = [];
  var notes = [];
  console.log(approvers);
  for (let approver of approvers)
  {
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
  for (let addId of adds['filesId'])
  {
    docs.push(DriveApp.getFileById(addId));
  }
  var comment = notes.join(";                 ");
  if (responses.indexOf(0) != -1)
  {
    var formName = FormApp.openByUrl(ss.getFormUrl()).getTitle();
    responseToRespondent(getRespondentEmail(ss, row), "Форма отклонена", formName, docs, comment);
  }
  else if (responses.length == 1 && responses.indexOf(1) != -1)
  {
    var formName = FormApp.openByUrl(ss.getFormUrl()).getTitle();
    responseToRespondent(getRespondentEmail(ss, row), "Форма одобрена", formName, docs, comment);
  }
}

function getRespondentEmail(ss, number)
{
  var sheet = ss.getActiveSheet();
  var lastCol = ss.getLastColumn();
  var range = sheet.getRange(1, 1, 1, lastCol);
  var values = range.getValues();
  for (var row in values) {
    for (var col in values[row]) {
      if (values[row][col] == "Адрес электронной почты")
      {
        return sheet.getRange(number, parseInt(col) + 1).getValue();
      }
    }
  }
  return "error";
}