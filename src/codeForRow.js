function quota() {      // функция для пользовательского меню, выводит на экран оставшийся для аккаунта лимит отправок писем
  var ui = SpreadsheetApp.getUi();
  var remaining = MailApp.getRemainingDailyQuota();
  ui.alert('📊 Оставшаяся квота на отправку писем',
           'На сегодня осталось: ' + remaining,
           ui.ButtonSet.OK);
}


function onOpen() {     // пользовательское меню
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📌 Отправка и создание по № строки')
    .addItem('📧 Отправить все документы ', 'sendAllPrompt')
    .addSeparator()         // линия разделитель
    .addItem('Оферта ➜ клиенту', 'sendOfferPrompt')
    .addSeparator()
    .addItem('Приложение ➜ менеджеру', 'sendAppPrompt')
    .addItem('Согласие ➜ менеджеру', 'sendConsentPrompt')
    .addSeparator()
    .addSeparator()
    .addItem('🔄 Создать документы заново', 'createDocuments')
    .addSeparator()
    .addSeparator()
    .addItem('💡 Лимит отправки', 'quota')
    .addToUi();
}
// clientData[0][9] - почта ЮЛ, clientData[0][6] - наименование ЮЛ
// clientData[0][17] - почта ФЛ, clientData[0][14] - ФИО ФЛ
// row - номер строки


function sendOfferPrompt() {        // функция для пользовательского меню отправки оферты клиенту из выбранной строки
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Отправка оферты', '\nВведите номер строки:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var row = parseInt(response.getResponseText(), 10);
    if (isNaN(row) || row < 1) return ui.alert('Ошибка', 'Введите корректный номер строки (целое число ≥1).', ui.ButtonSet.OK);
    try {
      sendPdfForRow(row);
      ui.alert('Успех', 'Оферта из строки ' + row + ' отправлена на почту клиента. \nЯчейка статуса закрашена зелёным.', ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Ошибка', e.message, ui.ButtonSet.OK);
    }
  }
}

function sendAppPrompt() {      // функция пользовательского меню отправки приложения менеджеру из выбранной строки
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Отправка приложения', '\nВведите номер строки:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var row = parseInt(response.getResponseText(), 10);
    if (isNaN(row) || row < 1) return ui.alert('Ошибка', 'Введите корректный номер строки.', ui.ButtonSet.OK);
    try {
      sendPdfdobForRow(row);
      ui.alert('Успех', 'Приложение из строки ' + row + ' отправлено на почту менеджера. \nЯчейка статуса закрашена зелёным.', ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Ошибка', e.message, ui.ButtonSet.OK);
    }
  }
}

function sendConsentPrompt() {      // функция пользовательского меню отправки согласия на ОПД менеджеру из выбранной строки
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Отправка согласия', '\nВведите номер строки:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var row = parseInt(response.getResponseText(), 10);
    if (isNaN(row) || row < 1) return ui.alert('Ошибка', 'Введите корректный номер строки.', ui.ButtonSet.OK);
    try {
      sendPdfAGRForRow(row);
      ui.alert('Успех', 'Согласие из строки ' + row + ' отправлено на почту менеджеру. \nЯчейка статуса закрашена зелёным.', ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Ошибка', e.message, ui.ButtonSet.OK);
    }
  }
}

function sendAllPrompt() {      // функция пользовательского меню, отправляет все документы адресатам из выбранной строки
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Отправка всех документов', 'Введите номер строки для отправки оферты, приложения и согласия:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var row = parseInt(response.getResponseText(), 10);
    if (isNaN(row) || row < 1) return ui.alert('Ошибка', 'Введите корректный номер строки.', ui.ButtonSet.OK);
    var errors = [];
    try { sendPdfForRow(row); } catch (e) { errors.push('Оферта: ' + e.message); }
    try { sendPdfdobForRow(row); } catch (e) { errors.push('Приложение: ' + e.message); }
    try { sendPdfAGRForRow(row); } catch (e) { errors.push('Согласие: ' + e.message); }
    if (errors.length === 0) {
      ui.alert('Успех', 'Все три документа для строки ' + row + ' отправлены. Ячейки статусов закрашены зелёным.', ui.ButtonSet.OK);
    } else {
      ui.alert('Частичный успех / ошибки', errors.join('\n'), ui.ButtonSet.OK);
    }
  }
}




//@param {number} rowNumber - номер строки в таблице (начиная с 1)

function sendPdfForRow(rowNumber) {     // преобразует оферту в PDF и отправляет письмо клиенту
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheets()[0];
  var lastR = sheet.getLastRow();
  
  if (rowNumber < 1 || rowNumber > lastR) throw new Error('Некорректный номер строки: ' + rowNumber);
  
  var clientData = sheet.getRange(rowNumber, 1, 1, 46).getValues();
  var id = clientData[0][41]; // ID оферты (столбец 42)
  
  if (!id) throw new Error('ID документа не найден в строке ' + rowNumber);
  
  var sendTo = [];
  if (clientData[0][9]) sendTo.push(clientData[0][9]);  // почта ЮЛ
  if (clientData[0][17]) sendTo.push(clientData[0][17]); // почта ФЛ
  
  if (sendTo.length === 0) throw new Error('Нет адресов электронной почты для отправки');
  
  var filepdf = DriveApp.getFileById(id).getAs('application/pdf');
  GmailApp.sendEmail(sendTo, clientData[0][6] + clientData[0][14] + '. Публичная оферта.',
    'Здравствуйте, ' + clientData[0][6] + clientData[0][14] + '.\n' +
    'Вы подключили свой автомобиль на реагирование вневедомственной охраной Росгвардии.\n\n' +
    'Вам необходимо направить копию свидетельства о регистрации ТС на почту gw@myavo.ru в наш клиентский отдел.\n\n' +
    'Тревожные сигналы отправляются в следующих случаях:\n\n' +
    '1. Вскрытие ТС – охранная система активирована, брелок иммобилайзера вне зоны системы (около 50 метров);\n' +
    '2. Нажатие кнопки «SOS» – оператору сообщено «ждем реагирования Росгвардией».\n\n' +
    'Круглосуточная техническая поддержка: 8 800 234 19 56\n\n' +
    'Во вложении - оферта, акцептованная Вами ранее.\n\n' +
    'Отвечать на это письмо не требуется.\n\nС уважением, команда ООО «Бизнес Мониторинг»',
    { attachments: [filepdf], name: 'Публичная оферта к договору на подключение системы T-Box GWM' });
  
  // Закрашиваем ячейку статуса (столбец 40) в зелёный цвет, не меняя значение
  sheet.getRange(rowNumber, 40).setBackground('#93c47d');
  sheet.getRange(rowNumber, 40).setValue('Отправлено');
  SpreadsheetApp.flush();
  console.log('Оферта для строки ' + rowNumber + ' отправлена, ячейка статуса закрашена зелёным.');
}


function sendPdfdobForRow(rowNumber) {      // преобразует приложение в PDF и отправляет письмо менеджеру
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheets()[0];
  var lastR = sheet.getLastRow();
  if (rowNumber < 1 || rowNumber > lastR) throw new Error('Некорректный номер строки: ' + rowNumber);
  
  var clientData = sheet.getRange(rowNumber, 1, 1, 46).getValues();
  var id = clientData[0][42]; // ID приложения (столбец 43)
  
  if (!id) throw new Error('ID приложения не найден в строке ' + rowNumber);
  
  var sendTo = clientData[0][2]; // почта менеджера
  if (!sendTo) throw new Error('Нет адреса менеджера для отправки');
  
  var filepdf = DriveApp.getFileById(id).getAs('application/pdf');
  GmailApp.sendEmail(sendTo, 'Приложение готово. Клиент: ' + clientData[0][6] + clientData[0][14],
    'Смотрите вложенный файл', { attachments: [filepdf], name: 'Приложение к публичной оферте' });
  
  // Закрашиваем ячейку статуса (столбец 44)
  sheet.getRange(rowNumber, 44).setBackground('#93c47d');
  sheet.getRange(rowNumber, 44).setValue('Отправлено');
  SpreadsheetApp.flush();
  console.log('Приложение для строки ' + rowNumber + ' отправлено, ячейка статуса закрашена зелёным.');
}


function sendPdfAGRForRow(rowNumber) {      // преобразует согласие в PDF и отправляет менеджеру
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheets()[0];
  var lastR = sheet.getLastRow();
  if (rowNumber < 1 || rowNumber > lastR) throw new Error('Некорректный номер строки: ' + rowNumber);
  
  var clientData = sheet.getRange(rowNumber, 1, 1, 46).getValues();
  var id = clientData[0][44]; // ID согласия (столбец 45)
  
  if (!id) throw new Error('ID согласия не найден в строке ' + rowNumber);
  
  var sendTo = clientData[0][2]; // почта менеджера
  if (!sendTo) throw new Error('Нет адреса менеджера для отправки');
  
  var filepdf = DriveApp.getFileById(id).getAs('application/pdf');
  GmailApp.sendEmail(sendTo, 'Согласие на ОПД готово. Клиент: ' + clientData[0][6] + clientData[0][14],
    'Смотрите вложенный файл', { attachments: [filepdf], name: 'Согласие на обработку персональных данных' });
  
  // Закрашиваем ячейку статуса (столбец 46)
  sheet.getRange(rowNumber, 46).setBackground('#93c47d');
  sheet.getRange(rowNumber, 46).setValue('Отправлено');
  SpreadsheetApp.flush();
  console.log('Согласие для строки ' + rowNumber + ' отправлено, ячейка статуса закрашена зелёным.');
}

function createDocuments() {        // функция пользовательского меню создания документов по номеру строки
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Создание документов', '\nВведите номер строки:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var row = parseInt(response.getResponseText(), 10);
    if (isNaN(row) || row < 1) {
      ui.alert('Ошибка', 'Введите корректный номер строки (целое число ≥1).', ui.ButtonSet.OK);
      return;
    }
    try {
      fillDocumentsForRow(row);
      // fillDocumentsForRow сама покажет уведомление об успехе или предупреждение
    } catch (e) {
      ui.alert('Ошибка', e.message, ui.ButtonSet.OK);
    }
  }
}

function fillDocumentsForRow(row) {     // копирует шаблоны, сохраяет в отдельном каталоге, заменяет плейсхолдеры данными из таблицы
  console.log(`[START] fillDocumentsForRow для строки ${row}`);
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  var ui = SpreadsheetApp.getUi();

  if (row < 2 || row > lastRow) {
    console.log(`[ERROR] Строка ${row} не существует. Допустимо 2..${lastRow}`);
    ui.alert('Ошибка', 'Строка ' + row + ' не существует. Допустимый диапазон: 2..' + lastRow, ui.ButtonSet.OK);
    return;
  }
  console.log(`[OK] Строка ${row} валидна`);

  console.log(`[DATA] Чтение данных строки ${row} (46 столбцов)`);
  var clientData = sheet.getRange(row, 1, 1, 46).getValues()[0];
  console.log(`[DATA] Данные получены: Организация="${clientData[6]}", ФЛ="${clientData[14]}"`);

  var date = clientData[4];
  var dPasd = clientData[21];
  var dBORN = clientData[15];
  var dPasdj = clientData[31];

  
  var months = { 1:'января',2:'февраля',3:'марта',4:'апреля',5:'мая',6:'июня',7:'июля',8:'августа',9:'сентября',10:'октября',11:'ноября',12:'декабря' };
  var monthsDig = { 1:'01',2:'02',3:'03',4:'04',5:'05',6:'06',7:'07',8:'08',9:'09',10:'10',11:'11',12:'12' };

  var splited = date.toString().split(" ");
  var realdata = '«' + splited[2] + '» ' + months[date.getMonth()+1] + ' ' + date.getFullYear() + ' года';
  var agrDate = splited[2] + ' ' + months[date.getMonth()+1] + ' ' + date.getFullYear() + ' г.';
  var agrNum = genAgrNumRow(row) + ' от ' + agrDate;
  var agrNumTS = genAgrNumRow(row);
  console.log(`[DATE] realdata="${realdata}", agrNum="${agrNum}"`);

  var datePASDJ = (dPasdj && dPasdj !== '') ? dPasdj.getDate().toString().padStart(2,'0')+'.'+monthsDig[dPasdj.getMonth()+1]+'.'+dPasdj.getFullYear() : '';
  var datePASD = (dPasd && dPasd !== '') ? dPasd.getDate().toString().padStart(2,'0')+'.'+monthsDig[dPasd.getMonth()+1]+'.'+dPasd.getFullYear() : '';
  var dateBORN = (dBORN && dBORN !== '') ? dBORN.getDate().toString().padStart(2,'0')+'.'+monthsDig[dBORN.getMonth()+1]+'.'+dBORN.getFullYear() : '';

  var dirTemp = DriveApp.getFolderById('1NGKQCrpFJhuz88uHRlrnyozp9RqYHwxt');
  console.log(`[DRIVE] Папка для копий получена`);

  console.log(`[COPY] Создание копии договора (Оферта)...`);
  var template = DriveApp.getFileById('148M3Bhc_z9FcQjE89JjYjF8dsMErpZoIOU9gZ6MOyOw')
    .makeCopy(clientData[6] + clientData[14] + ', ' + clientData[22] + ' Оферта', dirTemp).getId();
  console.log(`[COPY] Договор скопирован, ID=${template}`);

  console.log(`[COPY] Создание копии приложения...`);
  var templateDob = DriveApp.getFileById('1J49bfkps1luQYnWg_6J9l8MX-X0Z6tSpw5PEYzUuiFE')
    .makeCopy(clientData[6] + clientData[14] + ', ' + clientData[22] + ' Приложение', dirTemp).getId();
  console.log(`[COPY] Приложение скопировано, ID=${templateDob}`);

  console.log(`[COPY] Создание копии согласия...`);
  var templateAGR = DriveApp.getFileById('1fc09X5eLQICo96ArTRYa53LSL0WhePS0kLdVMj3vTwQ')
    .makeCopy(clientData[6] + clientData[14] + ', ' + clientData[22] + ' Соглашение на ОПД', dirTemp).getId();
  console.log(`[COPY] Согласие скопировано, ID=${templateAGR}`);

  Utilities.sleep(500);
  console.log(`[SLEEP] Пауза 500 мс после копирования`);

  // Заполнение договора
  console.log(`[DOC1] Открытие документа договора (ID=${template})`);
  var doc = DocumentApp.openById(template);
  var docbody = doc.getBody();
  console.log(`[DOC1] Замена плейсхолдеров...`);
  docbody.replaceText('DEALER', clientData[1]);
  docbody.replaceText('DATA', realdata);
  docbody.replaceText('NUM', agrNum);
  docbody.replaceText('NUTS', agrNumTS);
  docbody.replaceText('ORGNAME', clientData[6]);
  docbody.replaceText('FIOJL', clientData[7]);
  docbody.replaceText('TELEFONJL', clientData[8]);
  docbody.replaceText('EMAILJL', clientData[9]);
  docbody.replaceText('JURADR', clientData[10]);
  docbody.replaceText('POCHADR', clientData[11]);
  docbody.replaceText('INN', clientData[12]);
  docbody.replaceText('KPP', clientData[13]);
  docbody.replaceText('RS', clientData[36]);
  docbody.replaceText('KS', clientData[37]);
  docbody.replaceText('BANK', clientData[38]);
  docbody.replaceText('FIOFL', clientData[14]);
  docbody.replaceText('BORN', dateBORN);
  docbody.replaceText('TELEFONFL', clientData[16]);
  docbody.replaceText('EMAILFL', clientData[17]);
  docbody.replaceText('REGADR', clientData[18]);
  docbody.replaceText('SERPAS', clientData[19]);
  docbody.replaceText('PASS', clientData[20]);
  docbody.replaceText('PASD', datePASD);
  docbody.replaceText('PASDEP', clientData[34]);
  docbody.replaceText('PASC', clientData[35]);
  docbody.replaceText('TSVIN', clientData[22]);
  docbody.replaceText('NTS', clientData[23]);
  docbody.replaceText('MARK', clientData[24]);
  docbody.replaceText('MODEL', clientData[25]);
  docbody.replaceText('YTS', clientData[26]);
  docbody.replaceText('COLOR', clientData[27]);
  doc.saveAndClose();
  console.log(`[DOC1] Договор сохранён и закрыт`);

  // Заполнение приложения
  console.log(`[DOC2] Открытие документа приложения (ID=${templateDob})`);
  var doc2 = DocumentApp.openById(templateDob);
  var docbodyDob = doc2.getBody();
  console.log(`[DOC2] Замена плейсхолдеров...`);
  docbodyDob.replaceText('DEALER', clientData[1]);
  docbodyDob.replaceText('DATA', realdata);
  docbodyDob.replaceText('NUM', agrNum);
  docbodyDob.replaceText('NUTS', agrNumTS);
  docbodyDob.replaceText('ORGNAME', clientData[6]);
  docbodyDob.replaceText('FIOJL', clientData[7]);
  docbodyDob.replaceText('TELEFONJL', clientData[8]);
  docbodyDob.replaceText('EMAILJL', clientData[9]);
  docbodyDob.replaceText('JURADR', clientData[10]);
  docbodyDob.replaceText('POCHADR', clientData[11]);
  docbodyDob.replaceText('INN', clientData[12]);
  docbodyDob.replaceText('KPP', clientData[13]);
  docbodyDob.replaceText('RS', clientData[36]);
  docbodyDob.replaceText('KS', clientData[37]);
  docbodyDob.replaceText('BANK', clientData[38]);
  docbodyDob.replaceText('FIOFL', clientData[14]);
  docbodyDob.replaceText('BORN', dateBORN);
  docbodyDob.replaceText('TELEFONFL', clientData[16]);
  docbodyDob.replaceText('EMAILFL', clientData[17]);
  docbodyDob.replaceText('REGADR', clientData[18]);
  docbodyDob.replaceText('SERPAS', clientData[19]);
  docbodyDob.replaceText('PASS', clientData[20]);
  docbodyDob.replaceText('PASD', datePASD);
  docbodyDob.replaceText('PASDEP', clientData[34]);
  docbodyDob.replaceText('PASC', clientData[35]);
  docbodyDob.replaceText('TSVIN', clientData[22]);
  docbodyDob.replaceText('NTS', clientData[23]);
  docbodyDob.replaceText('MARK', clientData[24]);
  docbodyDob.replaceText('MODEL', clientData[25]);
  docbodyDob.replaceText('YTS', clientData[26]);
  docbodyDob.replaceText('COLOR', clientData[27]);
  doc2.saveAndClose();
  console.log(`[DOC2] Приложение сохранено и закрыто`);

  // Заполнение согласия
  console.log(`[DOC3] Открытие документа согласия (ID=${templateAGR})`);
  var doc3 = DocumentApp.openById(templateAGR);
  var docbodyAGR = doc3.getBody();
  console.log(`[DOC3] Замена плейсхолдеров...`);
  docbodyAGR.replaceText('FIOFL', clientData[14]);
  docbodyAGR.replaceText('FIOJL', clientData[7]);
  docbodyAGR.replaceText('EMAILJL', clientData[9]);
  docbodyAGR.replaceText('TELEFONJL', clientData[8]);
  docbodyAGR.replaceText('SERPASJ', clientData[28]);
  docbodyAGR.replaceText('NPASJ', clientData[29]);
  docbodyAGR.replaceText('PASDEPJ', clientData[30]);
  docbodyAGR.replaceText('PASDJ', datePASDJ);
  docbodyAGR.replaceText('PASCJ', clientData[32]);
  docbodyAGR.replaceText('ADRREGJ', clientData[33]);
  docbodyAGR.replaceText('SERPAS', clientData[19]);
  docbodyAGR.replaceText('PASS', clientData[20]);
  docbodyAGR.replaceText('PASDEP', clientData[34]);
  docbodyAGR.replaceText('PASC', clientData[35]);
  docbodyAGR.replaceText('REGADR', clientData[18]);
  docbodyAGR.replaceText('PASD', datePASD);
  docbodyAGR.replaceText('TELEFONFL', clientData[16]);
  docbodyAGR.replaceText('EMAILFL', clientData[17]);
  docbodyAGR.replaceText('MARK', clientData[24]);
  docbodyAGR.replaceText('MODEL', clientData[25]);
  docbodyAGR.replaceText('TSVIN', clientData[22]);
  doc3.saveAndClose();
  console.log(`[DOC3] Согласие сохранено и закрыто`);

  // Сохраняем ID документов в таблицу
  console.log(`[SHEET] Запись ID и статусов в строку ${row}, столбцы 40-46`);
  sheet.getRange(row, 42).setValue(template);
  sheet.getRange(row, 42).setBackground('#FFEBCD');
  sheet.getRange(row, 40).setValue('Новая');
  sheet.getRange(row, 41).setValue(agrNum);
  sheet.getRange(row, 41).setBackground('#FFEBCD');
  sheet.getRange(row, 43).setValue(templateDob);
  sheet.getRange(row, 43).setBackground('#FFEBCD');
  sheet.getRange(row, 44).setValue('Новая');
  sheet.getRange(row, 45).setValue(templateAGR);
  sheet.getRange(row, 45).setBackground('#FFEBCD');
  sheet.getRange(row, 46).setValue('Новая');

  SpreadsheetApp.flush();
  console.log(`[FLUSH] Данные таблицы принудительно записаны`);
  //ui.alert('Готово', 'Документы для строки ' + row + ' созданы и заполнены.', ui.ButtonSet.OK);
  ss.toast('Документы для строки ' + row + ' созданы', 'Готово', 5);
  console.log(`[END] fillDocumentsForRow для строки ${row} успешно завершена`);
}

function genAgrNumRow(row) {      // возвращает номер оферты (№ не используется при создании документов)
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheets()[0];
  var values = sheet.getDataRange().getValues();
  var num = randomNumRow(row);
  
  for (var i = 1; i < values.length; i++) {
    if (values[i][41] !== '') {
      checkNum = values[i][41];
      if (checkNum && typeof checkNum === 'string') {
        splited = checkNum.split(" ");
        if (splited.length > 0) {
          tempNum = splited[0].split("-");
          if (tempNum.length > 0) {
            if (parseInt(tempNum[0],27) == num) {
              num = randomNumRow(row);
              i = 1;
              continue;
            }
          }
        }
      }
    }
  }
  return 'GWM' + num + '-26';
}

function randomNumRow(rowNum) {       // рандомайзер для номера оферты
  return Math.floor(Math.random()*(99999-rowNum+1))+rowNum;
}