  /**
   * Version 1.0.0
   * Скрипт для обновления названия компании на листе "Реестр" по ИНН.
   * Если изменение произошло в столбце I, функция вызывает guessCompanyName и обновляет столбец J.
   */

  // Замените на свой API-ключ из личного кабинета (https://dadata.ru/profile/#info)
  var API_KEY = "50d5e3b21dcebbe86db095c11a398f5823d77f05";
  var VERSION = "1.0.0";

  /**
   * Функция updateCompanyName вызывается при изменении данных в листе "Реестр".
   * Если изменение произошло в столбце I, функция вызывает guessCompanyName и обновляет столбец J.
   *
   * @param {Event} e - Событие изменения данных в Google Sheets.
   */
  function updateCompanyName(e) {
    // Проверяем, что событие содержит необходимые данные
    if (!e || !e.range || !e.source) {
      return;
    }

    var sheet = e.source.getSheetByName("Реестр");
    if (!sheet) {
      return;
    }

    var range = e.range;
    
    // Проверяем, что изменение произошло в столбце I и не в первой строке
    if (range.getColumn() === 9 && range.getRow() > 1) {
      var startRow = range.getRow();
      var endRow = startRow + range.getNumRows() - 1;

      for (var row = startRow; row <= endRow; row++) {
        var inn = sheet.getRange(row, 9).getValue();

        // Проверяем, что ячейка в столбце I не пустая
        if (inn) {
          // Получаем название компании по ИНН
          var companyName = guessCompanyName(inn);
          sheet.getRange(row, 10).setValue(companyName);
        } else {
          // Если ячейка в столбце I пустая, очищаем ячейку в столбце J
          sheet.getRange(row, 10).clearContent();
        }
      }
    }
  }

  /**
   * Функция для получения названия компании по ИНН.
   * Эта функция вызывается напрямую из скрипта.
   *
   * @param {string} inn - ИНН компании.
   * @return {string} - Название компании.
   */
  function guessCompanyName(inn) {
    // Проверка наличия API-ключа
    if (API_KEY === "ВАШ_API_КЛЮЧ") {
      return "API ключ не указан";
    }

    // URL API DaData для поиска организации по ИНН
    var url = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party";
    
    // Подготовка данных для отправки POST-запроса
    var payload = {
      query: inn
    };
    
    // Настройка параметров запроса
    var options = {
      method: "POST",
      contentType: "application/json",
      headers: {
        "Authorization": "Token " + API_KEY
      },
      payload: JSON.stringify(payload)
    };
    
    try {
      // Отправка запроса и получение ответа
      var response = UrlFetchApp.fetch(url, options);
      
      // Преобразование ответа в объект JavaScript
      var responseData = JSON.parse(response.getContentText());
      
      // Проверка наличия результатов
      if (responseData.suggestions && responseData.suggestions.length > 0) {
        // Если результаты найдены, возвращаем значение названия организации
        return responseData.suggestions[0].value;
      } else {
        // Если результаты не найдены, возвращаем сообщение об отсутствии названия
        return "Название не найдено";
      }
    } catch (error) {
      return "Ошибка при запросе к API: " + error.message;
    }
  }