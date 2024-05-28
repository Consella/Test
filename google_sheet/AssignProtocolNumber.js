  /**
   * Version 1.0.0
   * Скрипт для присвоения номера протокола на листе "Реестр".
   * Преобразует дату из столбца Q и код пользователя из столбца C в уникальный номер протокола и записывает его в столбец A.
   */
  
  function setProtocolNumber() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Реестр');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Лист "Реестр" не найден.');
      return;
    }
  
    const userRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('user');
    if (!userRange) {
      SpreadsheetApp.getUi().alert('Именованный диапазон "user" не найден.');
      return;
    }
  
    const userData = userRange.getValues();
    const userMap = new Map();
    for (let i = 0; i < userData.length; i++) {
      userMap.set(userData[i][0], userData[i][1]);
    }
  
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
  
    const output = [];
    for (let i = 1; i < data.length; i++) { // Начинаем с 1, чтобы пропустить заголовок
      const cValue = data[i][2]; // Столбец C (индекс 2)
      const qValue = data[i][16]; // Столбец Q (индекс 16)
  
      if (cValue === "" || qValue === "") {
        output.push([""]);
      } else {
        const dateValue = new Date(qValue);
        const moscowTimeZone = "Europe/Moscow";
        const formattedDate = Utilities.formatDate(dateValue, moscowTimeZone, "yyMMdd");
        const lookupValue = userMap.get(cValue) || "";
        output.push([parseInt(formattedDate + lookupValue)]);
      }
    }
  
    const outputRange = sheet.getRange(2, 1, output.length, 1); // Записываем результаты в столбец A (индекс 1)
    outputRange.setValues(output);
  }