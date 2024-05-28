  /**
   * Copy data from "Registry" sheet to "Archive" sheet.
   * Only copies data from "Registry" if there is data in column R.
   * Checks for duplicates in "Archive" sheet before copying.
   * Adds a note with the copy time to the first cell of the copied row in the "Archive" sheet.
   * Version: 2.0.0
   */
  function copyToArchive() {
    var registrySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Реестр");
    var archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Архив");
    var lastRowRegistry = registrySheet.getLastRow();
    var lastRowArchive = archiveSheet.getLastRow();
    
    // Получаем массив данных из столбцов S и R на листе Реестр
    var columnRRegistry = registrySheet.getRange("R1:R" + lastRowRegistry).getValues();
    
    // Проходим по каждому значению в столбце R на листе Реестр
    for (var i = 0; i < columnRRegistry.length; i++) {
      // Проверяем, содержит ли текущая ячейка столбца R на листе Реестр данные
      if (columnRRegistry[i][0] !== "") {
        // Проверяем, есть ли уже такое значение в столбце R на листе Архив
        var isDuplicate = archiveSheet.getRange("R:R").getValues().flat().indexOf(columnRRegistry[i][0]) !== -1;
        
        // Если это уникальное значение, копируем текущую строку из листа Реестр в лист Архив
        if (!isDuplicate) {
          var rowToCopy = registrySheet.getRange(i + 1, 1, 1, registrySheet.getLastColumn()).getValues();
          archiveSheet.appendRow(rowToCopy[0]);
          
          // Добавляем заметку (комментарий) с временем копирования к первой ячейке соответствующей строки в листе Архив
          var copiedTime = new Date();
          var formattedTime = Utilities.formatDate(copiedTime, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
          var firstCell = archiveSheet.getRange(lastRowArchive + 1, 1);
          firstCell.setValue(rowToCopy[0][0]); // Значение первой ячейки равно значению из первой ячейки строки в Реестре
          firstCell.setNote("Скопировано " + formattedTime); // Заметка к первой ячейке
        }
      }
    }
  }
  
  function onArchiveCopy(e) {
    var range = e.range;
    var sheet = range.getSheet();
    
    // Проверяем, что изменение произошло в нужной нам таблице
    if (sheet.getName() == "Реестр" && range.getColumn() == 18) { // 18 - номер столбца R
      // Здесь вызывайте вашу функцию, которую нужно выполнить при изменении в столбце R
      // Например, copyToArchive();
      copyToArchive(); 
    }
  }