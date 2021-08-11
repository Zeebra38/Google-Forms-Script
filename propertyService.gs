function saveSettings(scriptUrl = "https://script.google.com/macros/s/AKfycbxI19xiYqrIRrvVrZhKRvUrIuBPHLWYdHx_-pSyG3gwFrd5y87k3y6xN5wYmWpo33FS/exec") {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('scriptURL', scriptUrl);
}

function loadSettings()
{
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperties();
}
