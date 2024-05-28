  /**
   * Версия: 1.0
   * Статус: Рабочий
   * Описание: Скрипт для экспорта диапазона A1:AX70 в PDF файл с именем, состоящим из слова "Протокол", номера из ячейки AZ6 и фамилии из ячейки B20.
   */
  
  function exportRangeToPDF() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getRange("A1:AX70");
    var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var sheetId = sheet.getSheetId();
    
    // Получаем номер из ячейки AZ6 и значение из ячейки B20
    var protocolNumber = sheet.getRange("AZ6").getValue();
    var fullName = sheet.getRange("B20").getValue();
    
    // Извлекаем фамилию из полного имени
    var lastName = fullName.split(' ')[0];
    
    // Формируем имя файла
    var fileName = "Протокол_№" + protocolNumber + "_" + lastName + ".pdf";
    
    var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?";
    var params = {
      format: "pdf",
      size: "A4",
      portrait: true, // Портретный формат
      fitw: true,
      gridlines: false,
      printtitle: false,
      sheetnames: false,
      pagenum: "UNDEFINED",
      attachment: false,
      gid: sheetId,
      range: "A1:AX70",
      top_margin: "0.5",
      bottom_margin: "0.5",
      left_margin: "0.2",
      right_margin: "0",
      horizontal_alignment: "CENTER",
      vertical_alignment: "TOP"
    };
    
    var queryString = [];
    for (var param in params) {
      queryString.push(param + "=" + encodeURIComponent(params[param]));
    }
    url += queryString.join("&");
    
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: {
        "Authorization": "Bearer " + token
      }
    });
    
    var blob = response.getBlob().setName(fileName);
    var folder = DriveApp.getFolderById('14rhWiuL-R8KtX537Kb3aURj73w-t12J2'); // Используем указанную папку
    folder.createFile(blob);
    
    // Выводим сообщение пользователю
    var ui = SpreadsheetApp.getUi();
    ui.alert("Протокол создан");
  }