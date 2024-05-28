  /**
   * Заполняет данные по пользователям в столбцах D, E, F, G и H листа "Реестр" на основе VLOOKUP для столбца C.
   * Если значение в столбце C пустое, оставляет соответствующие ячейки пустыми.
   * Если VLOOKUP не находит значения, устанавливает "Данных нет".
   */
  function fillUserData() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Реестр");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("Лист \"Реестр\" не найден.");
      return;
    }

    const userRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("user");
    if (!userRange) {
      SpreadsheetApp.getUi().alert("Именованный диапазон \"user\" не найден.");
      return;
    }

    const userData = userRange.getValues();
    const userMap = new Map();
    for (let i = 0; i < userData.length; i++) {
      userMap.set(userData[i][0], userData[i].slice(2, 7)); // Сохраняем значения из столбцов 3, 4, 5, 6 и 7
    }

    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();

    const output = [];
    for (let i = 1; i < data.length; i++) { // Начинаем с 1, чтобы пропустить заголовок
      const cValue = data[i][2]; // Столбец C (индекс 2)

      if (cValue === "") {
        output.push(["", "", "", "", ""]);
      } else {
        const lookupValues = userMap.get(cValue) || ["Данных нет", "Данных нет", "Данных нет", "Данных нет", "Данных нет"];
        output.push(lookupValues);
      }
    }

    const outputRange = sheet.getRange(2, 4, output.length, 5); // Записываем результаты в столбцы D, E, F, G и H (индексы 4, 5, 6, 7 и 8)
    outputRange.setValues(output);
  }

  // /** 
  //  * Обработчик триггера для отслеживания изменений в столбце C листа "Реестр".
  //  */
  // function onEdit(e) {
  //   const sheet = e.source.getActiveSheet();
  //   const range = e.range;

  //   if (sheet.getName() === "Реестр" && range.getColumn() === 3) { // Проверяем, что изменился столбец C
  //     fillUserData();
  //   }
  // }