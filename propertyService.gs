function saveSettings() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('scriptURL', 'https://script.google.com/macros/s/AKfycbwj1w8j6HP23rsLLwiQ8GdBVHA2I79dTSm2vKJZfTY/dev');
}

function loadSettings()
{
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperties();
}
