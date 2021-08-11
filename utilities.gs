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

function createComment(ss, row) {
  // ss = SpreadsheetApp.openById("1XT6aHEvD9AZvar8ypWYEFtibHGk-s6Ojx_Nmv04iK-A");
  // row = 10;
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
  console.log("Message = ", message);
  return message;
}

function checkFinalStage(ss, row) {
  // ss = SpreadsheetApp.openById("1XT6aHEvD9AZvar8ypWYEFtibHGk-s6Ojx_Nmv04iK-A");
  // row = 10;
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

function addslashes( str ) {
    return (str + '').replace(/[\\"']/g, '\\$&').replace(/\u0000/g, '\\0');
}