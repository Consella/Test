  /**
   * Version 1.4
   * Скрипт для копирования данных с листа "Реестр" на лист "Выгрузка".
   * Условия копирования: Если в столбце B листа "Реестр" значение true.
   * Соответствие столбцов:
   * D -> A
   * E -> B
   * F -> C
   * C -> D
   * H -> H
   * I -> I
   * J -> J
   * Q -> M
   * A -> N
   * В столбец K листа "Выгрузка" всегда вставляем значение "4345376178".
   * В столбец L листа "Выгрузка" всегда вставляем значение "Общество с ограниченной ответственностью \"Лига Качества\".
   * В столбец P листа "Выгрузка" всегда вставляем значение TRUE.
   * В столбец Q листа "Выгрузка" добавляем результаты replaceTrainingProgramCode.
   * В столбец O листа "Выгрузка" добавляем соответствие из столбца B листа "Программы".
   */

  function copyDataExport() {
    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Реестр");
    const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Выгрузка");
    const programSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Программы");

    if (!sourceSheet || !targetSheet || !programSheet) {
      SpreadsheetApp.getUi().alert("Лист \"Реестр\", \"Выгрузка\" или \"Программы\" не найден.");
      return false;
    }

    // Очищаем старые данные на листе "Выгрузка", начиная со второй строки
    targetSheet.getRange(2, 1, targetSheet.getMaxRows() - 1, targetSheet.getMaxColumns()).clearContent();

    const sourceData = sourceSheet.getDataRange().getValues();
    const programData = programSheet.getDataRange().getValues(); // Получаем все данные с листа "Программы"
    const targetData = [];

    // Создаем объект для быстрого поиска по столбцу A листа "Программы"
    const programMap = {};
    for (let i = 0; i < programData.length; i++) {
      const programCode = programData[i][0];
      const programResult = programData[i][5];
      const programDescription = programData[i][1];
      programMap[programCode] = { result: programResult, description: programDescription };
    }

    for (let i = 1; i < sourceData.length; i++) { // Начинаем с 1, чтобы пропустить заголовок
      if (sourceData[i][1] === true) { // Проверяем, что в столбце B значение true
        const row = new Array(18).fill(""); // Создаем массив с 18 элементами, заполненными пустыми строками
        row[0] = sourceData[i][3]; // D -> A
        row[1] = sourceData[i][4]; // E -> B
        row[2] = sourceData[i][5]; // F -> C
        row[3] = sourceData[i][2]; // C -> D
        row[7] = sourceData[i][7]; // H -> H
        row[8] = sourceData[i][8]; // I -> I
        row[9] = sourceData[i][9]; // J -> J
        row[10] = "4345376178"; // Постоянное значение для столбца K
        row[11] = "Общество с ограниченной ответственностью \"Лига Качества\"; // Постоянное значение для столбца L
        row[12] = sourceData[i][16]; // Q -> M
        row[13] = sourceData[i][0]; // A -> N
        row[15] = true; // Постоянное значение для столбца P

        // Добавляем результаты replaceTrainingProgramCode в столбец Q
        let queryResult = sourceData[i][10]; // Столбец K (индекс 10)
        if (queryResult == 1214 || queryResult == 1219 || queryResult == 1224) {
          queryResult = 3204;
        }
        const programResult = programMap[queryResult] ? programMap[queryResult].result : null;
        const programDescription = programMap[queryResult] ? programMap[queryResult].description : null;
        Logger.log(`Row ${i + 1}: queryResult = ${queryResult}, programResult = ${programResult}, programDescription = ${programDescription}`);
        row[16] = programResult; // Результат replaceTrainingProgramCode -> Q
        row[14] = programDescription; // Описание программы -> O

        targetData.push(row);
      }
    }

    if (targetData.length > 0) {
      targetSheet.getRange(2, 1, targetData.length, targetData[0].length).setValues(targetData);
      return true;
    } else {
      SpreadsheetApp.getUi().alert("Выберите галочкой данные. Нет данных для копирования.");
      return false;
    }
  }

  // Функция замены кода программ по высоте.
  function replaceTrainingProgramCode() {
    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Реестр");
    const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Выгрузка");
    const programSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Программы");

    if (!sourceSheet || !targetSheet || !programSheet) {
      SpreadsheetApp.getUi().alert("Лист \"Реестр\", \"Выгрузка\" или \"Программы\" не найден.");
      return;
    }

    const sourceData = sourceSheet.getDataRange().getValues();
    const programData = programSheet.getDataRange().getValues(); // Получаем все данные с листа "Программы"
    const targetData = [];

    // Создаем объект для быстрого поиска по столбцу A листа "Программы"
    const programMap = {};
    for (let i = 0; i < programData.length; i++) {
      const programCode = programData[i][0];
      const programResult = programData[i][5];
      programMap[programCode] = programResult;
    }

    for (let i = 1; i < sourceData.length; i++) { // Начинаем с 1, чтобы пропустить заголовок
      if (sourceData[i][1] === true) { // Проверяем, что в столбце B значение true
        let queryResult = sourceData[i][10]; // Столбец K (индекс 10)
        if (queryResult == 1214 || queryResult == 1219 || queryResult == 1224) {
          queryResult = 3204;
        }
        const programResult = programMap[queryResult] || null;
        Logger.log(`Row ${i + 1}: queryResult = ${queryResult}, programResult = ${programResult}`);
        targetData.push([programResult]);
      }
    }

    if (targetData.length > 0) {
      targetSheet.getRange(2, 1, targetData.length, 1).setValues(targetData);
      SpreadsheetApp.getUi().alert("Данные успешно обработаны и скопированы.");
    } else {
      SpreadsheetApp.getUi().alert("Нет данных для обработки.");
    }
  }