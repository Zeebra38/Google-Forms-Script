class Approver
{
  constructor(email, name, column)
  {
    this.name = name;
    this.email = email;
    this.column = column;
  }
}
function test2()
{
  e = {
  namedValues: 
   { 'Подразделение': [ 'Лаборатория 1' ],
     'email заказчика': [ 'email заказчика' ],
     'Дополнительные условия': [ 'Дополнительные условия' ],
     '': [ '' ],
     'Наименования товаров': [ 'Наименования товаров\nНаименования товаров\nНаименования товаров\nНаименования товаров\nНаименования товаров' ],
     'Адрес электронной почты': [ 'milekag318@obxstorm.com' ],
     'e-mail или телефон исполнителя': [ 'e-mail или телефон исполнителя' ],
     'Исполнитель': [ 'Исполнитель' ],
     'Приложение': [ '' ],
     'Отметка времени': [ '03.08.2021 9:29:02' ],
     'Организация заказчик': [ 'ООО' ] },
  range: { columnEnd: 10, columnStart: 1, rowEnd: 25, rowStart: 25 },
  source: {},
  triggerUid: '1736862934028947985',
  values: 
   [ '03.08.2021 9:29:02',
     'Лаборатория 1',
     'ООО',
     '',
     'Дополнительные условия',
     'e-mail или телефон исполнителя',
     'email заказчика',
     'Наименования товаров\nНаименования товаров\nНаименования товаров\nНаименования товаров\nНаименования товаров',
     'Исполнитель',
     'milekag318@obxstorm.com',
     '' ] };
  onFormSubmit(e);
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

function sendEmailWithAttach(email, subj, message, file) {
  MailApp.sendEmail(
    {
      to: email,
      subject: subj,
      htmlBody: message,
      attachments: file
    }
  );
}

function test()
{
  var templ = HtmlService.createTemplateFromFile('approveForm');
  templ.row = 5;
  templ.column = 5;
  templ.sheetID = "1XT6aHEvD9AZvar8ypWYEFtibHGk-s6Ojx_Nmv04iK-A";
  var message = templ.evaluate().getContent();
  sendEmail("bars00011@gmail.com", "test", message);
}

function responseToRespondent(email, subject, formName, file, comment) {
  var templ;
  switch (subject)
  {
    case "Форма отклонена":
    templ = HtmlService.createTemplateFromFile('rejectedForm');
    break;
    case "Форма одобрена":
    templ = HtmlService.createTemplateFromFile('approvedForm');
    break;
    case "Форма принята к рассмотрению":
    templ = HtmlService.createTemplateFromFile('starterNotification');
    break;
  }
  templ.formName = formName;
  var message = templ.evaluate().getContent();
  // sendEmailWithAttach(email, subject, message, [file]);
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
  // row = 15;
  var sheet = ss.getActiveSheet();
  var name = ss.getName();
  var docFolderName = name.replace('(Ответы)', 'Документы');
  var curFile = DriveApp.getFileById(ss.getId());
  var folderId = curFile.getParents().next().getId();
  var dir = DriveApp.getFolderById(folderId);
  var docFolder = dir.getFoldersByName(docFolderName).next();
  var doc = docFolder.searchFiles(`title contains "Записка №${row}"`).next();
  getListOfApprovers(ss).forEach(function(approver) {
    sheet.getRange(row, approver.column).setValue("?");
    var templ = HtmlService.createTemplateFromFile('approveForm');
    templ.row = row;
    templ.column = approver.column;
    templ.ssID = ss.getId();
    var message = templ.evaluate().getContent();
    // sendEmailWithAttach(approver.email, "Подтвердите форму", message, [doc]);
  });
}

function readyCheck(ss, row, column)
{
  var sheet = ss.getActiveSheet();
  var lastCol = ss.getLastColumn();
  var approvers = getListOfApprovers(ss);
  var responses = [];
  console.log(approvers);
  for (let approver of approvers)
  {
    responses.push(sheet.getRange(row, approver.column).getValue());
  }
  responses = [...new Set(responses)];
  console.log(responses);
  var name = ss.getName();
  var docFolderName = name.replace('(Ответы)', 'Документы');
  var curFile = DriveApp.getFileById(ss.getId());
  var folderId = curFile.getParents().next().getId();
  var dir = DriveApp.getFolderById(folderId);
  var docFolder = dir.getFoldersByName(docFolderName).next();
  var doc = docFolder.searchFiles(`title contains "Записка №${row}"`).next();
  if (responses.indexOf(0) != -1)
  {
    var formName = FormApp.openByUrl(ss.getFormUrl()).getTitle();
    responseToRespondent(getRespondentEmail(ss, row), "Форма отклонена", formName, doc);
  }
  else if (responses.length == 1 && responses.indexOf(1) != -1)
  {
    var formName = FormApp.openByUrl(ss.getFormUrl()).getTitle();
    responseToRespondent(getRespondentEmail(ss, row), "Форма одобрена", formName, doc);
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