/**
 * Фукнция для клиента. Устанавливает триггеры на все Таблицы по их Id, разделенные '\n'
 */
function starter() {
  // Писать вот так. Использовать только Enter и писать Id
  var example = `123
  456aa
  asdsa23124
  dsdsa65fd-123z`; 
  var spreadsheetsIDs = `1XT6aHEvD9AZvar8ypWYEFtibHGk-s6Ojx_Nmv04iK-A
  `;
  mainSetup(spreadsheetsIDs);
}
