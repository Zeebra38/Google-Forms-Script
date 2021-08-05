function saveSettings() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('scriptURL', 'https://script.google.com/macros/s/AKfycbxI19xiYqrIRrvVrZhKRvUrIuBPHLWYdHx_-pSyG3gwFrd5y87k3y6xN5wYmWpo33FS/exec');
}

function loadSettings()
{
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperties();
}
