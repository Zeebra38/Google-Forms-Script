/**
 * Обработчик события Get
 * @params {event} e Событие, которое передается при вызове функции. Содержит ключ 'parameter', из которого формируется Object params. Словарь ключ-значение
 */
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
/**
 * Обработчик события Post
 */
function doPost(e) {
  console.log(e);
  var params = e['parameter'];
  console.log(params);
  return HtmlService.createHtmlOutput("OK");
}

/**
 * Генерирует HtmlOutput. Это страница является результатом, на которую попадает "одобритель" после принятия/отклонения формы
 * @params {object} params Словарь ключ-значение, содержащий данные
 */
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

/**
 * Генерирует HtmlOutput из шаблона approveForm. Это страница одобрения/отклонения формы
 * @params {object} params Словарь ключ-значение, содержащий данные
 */
function goToResponsePage(params) {
  var templ = HtmlService.createTemplateFromFile('approveForm');
  templ.row = params.row;
  templ.column = params.column;
  templ.ssID = params.ssID;
  templ.scriptURL = loadSettings().scriptURL;
  return templ.evaluate();
}