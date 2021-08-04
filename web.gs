class approveForm
{
  constructor(email, name, column)
  {
    this.name = name;
    this.email = email;
    this.column = column;
  }
}

function doGet(e) {
  console.log(e);
  var params = e['parameter'];
  console.log(params);
  var ss = SpreadsheetApp.openById(params['ssID']);
  var sheet = ss.getActiveSheet();
  var row = params['row'];
  var column = params['column'];
  var cell = sheet.getRange(row, column);
  if (params['handler'] == 'Approved')
  {
    cell.setValue("1");
    readyCheck(ss, row, column);
    return HtmlService.createHtmlOutput(`Вы подтвердили форму`);
  }
    if (params['handler'] == 'Rejected')
  {
    cell.setValue("0");
    readyCheck(ss, row, column);
    // return HtmlService.createHtmlOutput(`Вы отклонили форму`);
    return HtmlService.createHtmlOutputFromFile(`approveForm`);
  }
  return HtmlService.createHtmlOutput("Ошибка");
}

function doPost(e)
{
  var params = JSON.parse(e);
  return HtmlService.createHtmlOutput(params);
}