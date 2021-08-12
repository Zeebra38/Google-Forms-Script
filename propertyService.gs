/**
 * Сохраряент настройки для текущего скрипта
 * @param {string} scriptUrl Ссылка на развернутое Web-приложение
 */
function saveSettings(scriptUrl = "") {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('scriptURL', scriptUrl);
}

/**
 * Загружает настроки для текущего проекта
 */
function loadSettings()
{
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperties();
}
