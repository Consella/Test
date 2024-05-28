  /**
   * Version 1.4
   * Скрипт для экспорта данных с листа "Выгрузка" в XML файл.
   * Проверяет корректность даты перед добавлением в XML.
   * Если дата недействительна, элемент Date добавляется со значением по умолчанию.
   * Если learnProgramId недействителен, атрибут learnProgramId не добавляется.
   */
  
  function exportToXML() {
    const folderId = '1UVQbi3GxahHI-SwWBvVl4IVJCWshnOLd'; // ID папки
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Выгрузка');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Лист "Выгрузка" не найден.');
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    
    const xmlDoc = XmlService.createDocument();
    const registrySet = XmlService.createElement('RegistrySet');
    xmlDoc.setRootElement(registrySet);
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== "") { // Проверяем, что строка не пустая
        const registryRecord = XmlService.createElement('RegistryRecord');
        
        const worker = XmlService.createElement('Worker');
        worker.addContent(XmlService.createElement('LastName').setText(data[i][0]));
        worker.addContent(XmlService.createElement('FirstName').setText(data[i][1]));
        worker.addContent(XmlService.createElement('MiddleName').setText(data[i][2]));
        worker.addContent(XmlService.createElement('Snils').setText(data[i][3]));
        worker.addContent(XmlService.createElement('IsForeignSnils').setText(validateBit(data[i][4])));
        worker.addContent(XmlService.createElement('ForeignSnils').setText(data[i][5]));
        worker.addContent(XmlService.createElement('Citizenship').setText(data[i][6]));
        worker.addContent(XmlService.createElement('Position').setText(data[i][7]));
        worker.addContent(XmlService.createElement('EmployerInn').setText(data[i][8]));
        worker.addContent(XmlService.createElement('EmployerTitle').setText(data[i][9]));
        
        const organization = XmlService.createElement('Organization');
        organization.addContent(XmlService.createElement('Inn').setText(data[i][10]));
        organization.addContent(XmlService.createElement('Title').setText(data[i][11]));
        
        const test = XmlService.createElement('Test');
        test.setAttribute('isPassed', data[i][15]);
        
        const learnProgramId = data[i][16];
        if (learnProgramId) {
          test.setAttribute('learnProgramId', learnProgramId);
        }
        
        const formattedDate = formatDate(data[i][12]);
        if (formattedDate) {
          test.addContent(XmlService.createElement('Date').setText(formattedDate));
        } else {
          test.addContent(XmlService.createElement('Date').setText('1900-01-01')); // Добавляем значение по умолчанию для недействительных дат
        }
        test.addContent(XmlService.createElement('ProtocolNumber').setText(data[i][13]));
        test.addContent(XmlService.createElement('LearnProgramTitle').setText(data[i][14]));
        
        registryRecord.addContent(worker);
        registryRecord.addContent(organization);
        registryRecord.addContent(test);
        
        registrySet.addContent(registryRecord);
      }
    }
    
    const xmlOutput = XmlService.getPrettyFormat().format(xmlDoc);
    const timestamp = getFormattedTimestamp();
    const fileName = `RegistrySet_${timestamp}.xml`;
    const blob = Utilities.newBlob(xmlOutput, 'application/xml', fileName);
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    
    Logger.log('XML file created: '+file.getUrl());
    SpreadsheetApp.getUi().alert('Файл сформирован.');
  }
  
  function formatDate(date) {
    const d = new Date(date);
    if (isNaN(d.getTime())) {
      return null; // Возвращаем null, если дата недействительна
    }
    const year = d.getFullYear();
    const month = ('0' + (d.getMonth() + 1)).slice(-2);
    const day = ('0' + d.getDate()).slice(-2);
    return `${year}-${month}-${day}`;
  }
  
  function getFormattedTimestamp() {
    const now = new Date();
    const timezone = 'GMT+3'; // Укажите ваш часовой пояс
    const formattedDate = Utilities.formatDate(now, timezone, 'dd-MM-yyyy_HH-mm');
    return formattedDate;
  }
  
  function validateBit(value) {
    const validBits = ['0', '1', 'False', 'True', 'false', 'true', 'FALSE', 'TRUE'];
    return validBits.includes(value) ? value : 'False';
  }