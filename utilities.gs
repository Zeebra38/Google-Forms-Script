/**
 * Формирует тему письма
 * @param {Spreadsheet} ss Гугл Таблица
 * @param {number} row Номер строки
 */
function getSubject(ss, row) {
  var name = ss.getName();
  var docFolderName = name.replace('(Ответы)', 'Документы');
  var curFile = DriveApp.getFileById(ss.getId());
  var folderId = curFile.getParents().next().getId();
  var dir = DriveApp.getFolderById(folderId);
  var docFolder = dir.getFoldersByName(docFolderName).next();
  var doc = docFolder.searchFiles(`title contains "Записка №${row}"`).next();
  return doc.getName().slice(0, -4) + " ";
}

/**
 * Находит комментарии к ответам 
 * @param {Spreadsheet} ss Гугл Таблица
 * @param {number} row Номер строки
 */
function createComment(ss, row) {
  var sheet = ss.getActiveSheet();
  var approvers = getListOfApprovers(ss);
  var comment = "";
  for (var i = 0; i < approvers.length-1; i++) {
    if (sheet.getRange(row, approvers[i].column).getNote() != undefined && sheet.getRange(row, approvers[i].column).getNote() != "")
    {
      comment += `${approvers[i].email}: ${sheet.getRange(row, approvers[i].column).getNote()}<br>`;
    }
    else
    {
      comment += `${approvers[i].email}: без комментариев <br>`;
    }
  }
  var message = HtmlService.createHtmlOutput(comment).getContent();
  return message;
}

/**
 * Возвращает true, если документ успешно прошел одобрение всеми и должен быть отправлен Адресату
 * @param {Spreadsheet} ss Гугл Таблица
 * @param {number} row Номер строки
 */
function checkFinalStage(ss, row) {
  var sheet = ss.getActiveSheet();
  var approvers = getListOfApprovers(ss);
  for (var i = 0; i < approvers.length - 1; i++) {
    if(sheet.getRange(row, approvers[i].column).getValue() != "Подтверждено")
    {
      return false;
    }
  }
  return true;
}

/**
 * Возвращает список email одобрителей в виде строки, которая потом используется для вставки в контент htmlTemplate
 * @param {Spreadsheet} ss Гугл Таблица
 */
function getApproversToEmail(ss)
{
  var approvers = getListOfApprovers(ss);
  var approversEmails = [];
  for(var i = 0; i < approvers.length - 1; i++)
  {
    approversEmails.push(approvers[i].email);
  }
  return approversEmails.join(", ");
}