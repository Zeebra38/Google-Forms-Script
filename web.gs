function doGet(e) {
  console.log(e);
  var params = e['parameter'];
  console.log(params);
  switch (params.action) {
    case "sendResponse":
      return sendResponse(params);
    case "goToResponsePage":
      return goToResponsePage(params);
    default:
      return HtmlService.createHtmlOutput("Ошибка 500");
  }
}

function doPost(e) {
  console.log(e);
  var params = e['parameter'];
  console.log(params);
  return HtmlService.createHtmlOutput("OK");
}

function sendResponse(params) {
  var ss = SpreadsheetApp.openById(params['ssID']);
  var sheet = ss.getActiveSheet();
  var row = params['row'];
  var column = params['column'];
  var cell = sheet.getRange(row, column);
  if (cell.getValue() == "На обработке" || cell.getValue() == "Отклонено") {
    cell.clearNote();
    switch (params.handler) {
      case 'Approved':
        cell.setValue("Подтверждено");
        cell.setNote(params.comment.replaceAll("+", " "));
        readyCheck(ss, row, column);
        return HtmlService.createHtmlOutput(`Вы успешно подтвердили форму`);
      case 'Rejected':
        cell.setValue("Отклонено");
        cell.setNote(params.comment.replaceAll("+", " "));
        readyCheck(ss, row, column);
        return HtmlService.createHtmlOutput(`Вы успешно отклонили форму`);
      default:
        return HtmlService.createHtmlOutput("Ошибка 500");
    }
  }
  else {
    return HtmlService.createHtmlOutput('Ошибка. Вы уже одобрили форму.')
  }
}

function goToResponsePage(params) {
  var templ = HtmlService.createTemplateFromFile('approveForm');
  templ.row = params.row;
  templ.column = params.column;
  templ.ssID = params.ssID;
  templ.scriptURL = loadSettings().scriptURL;
  return templ.evaluate();
}