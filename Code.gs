/**
 * ЗАЯВКА IT В БЮДЖЕТ MTK — backend бота (Google Apps Script)
 * ------------------------------------------------------------
 * Как развернуть — см. README.md. Коротко:
 *  1. Скрипт должен быть привязан к Google-таблице (Расширения → Apps Script).
 *  2. Заполнить SCRIPT PROPERTIES: BOT_TOKEN (токен от @BotFather).
 *  3. Выполнить setup() один раз — создаст листы "Каталог", "Заявки", "Позиции".
 *  4. Развернуть как веб-приложение (Deploy → New deployment → Web app),
 *     доступ "Anyone". Скопировать exec-URL и сохранить его в свойстве
 *     WEBAPP_URL (без этого сработает менее надёжное автоопределение адреса).
 *     Веб-приложение нужно только для открытия формы (Mini App) из Telegram.
 *  5. Бот получает сообщения через ОПРОС (polling), а не через webhook:
 *     Google Apps Script веб-приложения всегда отвечают редиректом 302 на
 *     запрос к /exec, а Telegram не следует за редиректами при доставке
 *     вебхуков — поэтому setWebhook() для этого бэкенда не работает.
 *     Выполните startPolling() один раз — она удалит вебхук (если был) и
 *     создаст триггер по времени, который каждую минуту вызывает
 *     pollUpdates() и забирает новые сообщения через getUpdates().
 */

// ==================== НАСТРОЙКИ ====================

// Фиксированный список подразделений: код -> отображаемое имя.
// Код используется в callback_data (ограничение Telegram — 64 байта),
// поэтому в кнопках всегда только эти 7 пунктов, добавить новый нельзя.
var DEPARTMENTS = {
  D1: 'Отдел закупа',
  D2: 'Юридическая группа',
  D3: 'Группа бюджетирования',
  D4: 'Административно-хозяйственная группа',
  D5: 'Производственно-технический отдел',
  D6: 'Сектор производственных отношений',
  D7: 'Отдел кадров'
};

// Единицы измерения — по умолчанию первая ("шт."), остальные доступны в списке.
var UNIT_OPTIONS = ['шт.', 'компл.', 'уп.', 'пара', 'м', 'кг', 'лицензия', 'коробка'];

// Чек-лист "возможно забытого" оборудования/периферии — предлагается перед
// отправкой заявки, чтобы не упустить типовые позиции при планировании
// бюджета на СЛЕДУЮЩИЙ год.
var FORGOTTEN_CHECKLIST = [
  'Компьютерная мышь',
  'Клавиатура',
  'Гарнитура (наушники с микрофоном)',
  'Веб-камера',
  'Дополнительный монитор',
  'ИБП (источник бесперебойного питания)',
  'USB-хаб / докинг-станция',
  'Сетевой кабель (патч-корд)',
  'Сетевой фильтр / удлинитель',
  'Внешний накопитель / флеш-память',
  'Картридж или тонер для принтера',
  'МФУ / принтер',
  'Лицензия на программное обеспечение'
];

var MAX_ITEMS_PER_REQUEST = 20;
var JUSTIFICATION_MIN_LENGTH = 8;

var SHEET_CATALOG = 'Каталог';
var SHEET_REQUESTS = 'Заявки';
var SHEET_ITEMS = 'Позиции';

function getBudgetYear() {
  var override = PropertiesService.getScriptProperties().getProperty('BUDGET_YEAR');
  if (override) return parseInt(override, 10);
  return new Date().getFullYear() + 1; // по умолчанию — следующий год
}

function getBotToken() {
  var token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  if (!token) throw new Error('BOT_TOKEN не задан в Script Properties.');
  return token;
}

function getWebAppUrl() {
  var stored = PropertiesService.getScriptProperties().getProperty('WEBAPP_URL');
  return stored || ScriptApp.getService().getUrl();
}

// ==================== ОДНОРАЗОВАЯ НАСТРОЙКА ====================

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var catalog = ss.getSheetByName(SHEET_CATALOG);
  if (!catalog) {
    catalog = ss.insertSheet(SHEET_CATALOG);
    catalog.getRange(1, 1, 1, 2).setValues([['Наименование товара', 'Категория (необязательно)']]);
    catalog.getRange(2, 1, 3, 2).setValues([
      ['Ноутбук бизнес-класса 14"', 'Компьютерная техника'],
      ['Монитор 24"', 'Периферия'],
      ['МФУ лазерное А4', 'Оргтехника']
    ]);
    catalog.setFrozenRows(1);
  }

  var requests = ss.getSheetByName(SHEET_REQUESTS);
  if (!requests) {
    requests = ss.insertSheet(SHEET_REQUESTS);
    requests.getRange(1, 1, 1, 8).setValues([[
      'ID заявки', 'Отдел', 'Год бюджета', 'Статус',
      'Создана', 'Изменена', 'Кол-во позиций', 'Chat ID автора'
    ]]);
    requests.setFrozenRows(1);
  }

  var items = ss.getSheetByName(SHEET_ITEMS);
  if (!items) {
    items = ss.insertSheet(SHEET_ITEMS);
    items.getRange(1, 1, 1, 8).setValues([[
      'ID заявки', '№', 'Наименование', 'Есть в каталоге',
      'Обоснование', 'Количество', 'Ед. изм.', 'Ссылка'
    ]]);
    items.setFrozenRows(1);
  }

  SpreadsheetApp.flush();
  Logger.log('Листы созданы/проверены. Заполните "Каталог" реальными позициями IT-каталога.');
}

// ==================== ПОЛУЧЕНИЕ СООБЩЕНИЙ (POLLING) ====================
//
// Веб-приложения Google Apps Script всегда отвечают HTTP 302 (редирект на
// script.googleusercontent.com) на запрос к /exec. Браузеры и большинство
// HTTP-клиентов следуют за таким редиректом автоматически, но Telegram при
// доставке webhook-обновлений редиректы не проверяет и считает 302 ошибкой
// доставки — поэтому setWebhook() для этого бэкенда не работает. Вместо
// этого бот сам опрашивает Telegram через getUpdates() по триггеру времени.

function startPolling() {
  // на всякий случай снимаем вебхук — getUpdates() и webhook несовместимы
  UrlFetchApp.fetch('https://api.telegram.org/bot' + getBotToken() + '/deleteWebhook', { muteHttpExceptions: true });

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pollUpdates') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pollUpdates').timeBased().everyMinutes(1).create();

  Logger.log('Опрос запущен: pollUpdates() будет выполняться каждую минуту.');
}

function stopPolling() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pollUpdates') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Триггер опроса удалён.');
}

function pollUpdates() {
  var props = PropertiesService.getScriptProperties();
  var offset = Number(props.getProperty('LAST_UPDATE_ID') || 0) + 1;

  var resp = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + getBotToken() + '/getUpdates?offset=' + offset + '&timeout=0',
    { muteHttpExceptions: true }
  );
  var data = JSON.parse(resp.getContentText());
  if (!data.ok || !data.result.length) return;

  data.result.forEach(function (update) {
    try {
      handleTelegramUpdate(update);
    } finally {
      props.setProperty('LAST_UPDATE_ID', String(update.update_id));
    }
  });
}

// ==================== ВХОДЯЩИЕ ЗАПРОСЫ ВЕБ-ПРИЛОЖЕНИЯ ====================

function doGet(e) {
  var dept = DEPARTMENTS[e.parameter.dept];
  if (!dept) {
    return HtmlService.createHtmlOutput('Не указано подразделение. Откройте форму из бота.');
  }
  var year = e.parameter.year || getBudgetYear();
  var reqId = e.parameter.reqId || '';
  var chatId = e.parameter.chatId || '';

  var existingItems = reqId ? getItemsForRequest(reqId) : [];
  var remaining = MAX_ITEMS_PER_REQUEST - existingItems.length;

  var tpl = HtmlService.createTemplateFromFile('Form');
  tpl.department = dept;
  tpl.deptCode = e.parameter.dept;
  tpl.year = year;
  tpl.reqId = reqId;
  tpl.chatId = chatId;
  tpl.remaining = remaining;
  tpl.catalogJson = JSON.stringify(getCatalog());
  tpl.unitsJson = JSON.stringify(UNIT_OPTIONS);
  tpl.checklistJson = JSON.stringify(FORGOTTEN_CHECKLIST);
  tpl.existingItemsJson = JSON.stringify(existingItems);
  tpl.justificationMin = JUSTIFICATION_MIN_LENGTH;
  tpl.postUrl = getWebAppUrl();

  return tpl.evaluate()
    .setTitle('Заявка IT в бюджет MTK')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput('bad request');
  }

  if (body.miniapp) {
    handleMiniAppSubmit(body);
  } else if (body.message || body.callback_query) {
    handleTelegramUpdate(body);
  }
  return ContentService.createTextOutput('ok');
}

// ==================== ЛОГИКА TELEGRAM-БОТА ====================

function handleTelegramUpdate(update) {
  if (update.message) {
    var msg = update.message;
    var chatId = msg.chat.id;
    var text = (msg.text || '').trim();

    if (text === '/start' || text === '/newrequest' || text === 'Новая заявка') {
      sendMessage(chatId,
        'Здравствуйте! 👋\n\n' +
        'Это бот «Заявка IT в бюджет MTK» — здесь собираются заявки на офисное ' +
        'IT-оборудование при формировании бюджета на ' + getBudgetYear() + ' год.\n\n' +
        'Выберите ваше структурное подразделение:',
        departmentKeyboard()
      );
    } else {
      sendMessage(chatId, 'Чтобы начать заявку, отправьте команду /start.');
    }
    return;
  }

  if (update.callback_query) {
    var cq = update.callback_query;
    var chatId2 = cq.message.chat.id;
    var data = cq.data || '';
    answerCallbackQuery(cq.id);

    if (data.indexOf('dept:') === 0) {
      var deptCode = data.split(':')[1];
      openDepartmentFlow(chatId2, deptCode);
    } else if (data.indexOf('continue:') === 0) {
      var parts = data.split(':'); // continue:<deptCode>:<reqId>
      openWebApp(chatId2, parts[1], parts[2]);
    } else if (data.indexOf('view:') === 0) {
      var reqId2 = data.split(':')[1];
      sendRequestSummary(chatId2, reqId2);
    }
    return;
  }
}

function departmentKeyboard() {
  var rows = [];
  Object.keys(DEPARTMENTS).forEach(function (code) {
    rows.push([{ text: DEPARTMENTS[code], callback_data: 'dept:' + code }]);
  });
  return { inline_keyboard: rows };
}

function openDepartmentFlow(chatId, deptCode) {
  var dept = DEPARTMENTS[deptCode];
  var year = getBudgetYear();
  var existing = findActiveRequest(dept, year);

  if (existing) {
    sendMessage(chatId,
      'У подразделения «' + dept + '» уже есть активная заявка на ' + year + ' год.\n' +
      'Номер: ' + existing.id + '\n' +
      'Позиций внесено: ' + existing.count + ' из ' + MAX_ITEMS_PER_REQUEST + '\n' +
      'Последнее изменение: ' + Utilities.formatDate(existing.updated, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') + '\n\n' +
      'Чтобы не создавать вторую заявку по этому же подразделению, дополните существующую.',
      {
        inline_keyboard: [
          [{ text: '➕ Дополнить заявку', callback_data: 'continue:' + deptCode + ':' + existing.id }],
          [{ text: '📋 Показать список позиций', callback_data: 'view:' + existing.id }]
        ]
      }
    );
  } else {
    openWebApp(chatId, deptCode, '');
  }
}

function openWebApp(chatId, deptCode, reqId) {
  var year = getBudgetYear();
  var url = getWebAppUrl() +
    '?dept=' + encodeURIComponent(deptCode) +
    '&year=' + encodeURIComponent(year) +
    '&chatId=' + encodeURIComponent(chatId) +
    '&reqId=' + encodeURIComponent(reqId || '');

  sendMessage(chatId,
    reqId ? 'Открываю заявку для дополнения:' : 'Открываю форму новой заявки:',
    { inline_keyboard: [[{ text: '🧾 Открыть форму', web_app: { url: url } }]] }
  );
}

function sendRequestSummary(chatId, reqId) {
  var items = getItemsForRequest(reqId);
  if (!items.length) {
    sendMessage(chatId, 'В заявке ' + reqId + ' пока нет позиций.');
    return;
  }
  var lines = items.map(function (it, i) {
    return (i + 1) + '. ' + it.name + ' — ' + it.qty + ' ' + it.unit;
  });
  sendMessage(chatId, 'Заявка ' + reqId + ', внесённые позиции:\n\n' + lines.join('\n'));
}

// ==================== ПРИЁМ ФОРМЫ ИЗ MINI APP ====================

function handleMiniAppSubmit(body) {
  var deptCode = body.dept;
  var dept = DEPARTMENTS[deptCode];
  var chatId = body.chatId;
  var year = body.year || getBudgetYear();

  if (!dept || !chatId) {
    return; // некорректные данные, молча игнорируем
  }

  var reqId = body.reqId;
  var isNew = !reqId;
  if (isNew) {
    reqId = generateRequestId(deptCode, year);
    addRequestRow(reqId, dept, year, chatId);
  }

  var existingCount = getItemsForRequest(reqId).length;
  var slotsLeft = MAX_ITEMS_PER_REQUEST - existingCount;

  var catalog = getCatalog();
  var accepted = [];
  var rejected = [];

  (body.items || []).forEach(function (item) {
    if (accepted.length >= slotsLeft) {
      rejected.push({ name: item.name, reason: 'превышен лимit ' + MAX_ITEMS_PER_REQUEST + ' позиций в заявке' });
      return;
    }
    var name = (item.name || '').trim();
    var justification = (item.justification || '').trim();
    var qty = parseInt(item.qty, 10);
    var unit = item.unit || 'шт.';
    var link = (item.link || '').trim();

    if (!name || !justification || !qty || qty < 1) {
      rejected.push({ name: name || '(без названия)', reason: 'не заполнены обязательные поля' });
      return;
    }
    if (justification.length < JUSTIFICATION_MIN_LENGTH) {
      rejected.push({ name: name, reason: 'обоснование короче ' + JUSTIFICATION_MIN_LENGTH + ' символов' });
      return;
    }
    var inCatalog = catalog.some(function (c) { return c.toLowerCase() === name.toLowerCase(); });
    accepted.push({ name: name, inCatalog: inCatalog, justification: justification, qty: qty, unit: unit, link: link });
  });

  if (accepted.length) {
    appendItems(reqId, accepted);
    updateRequestSummary(reqId);
  }

  var text = '';
  if (isNew && accepted.length) {
    text += 'Заявка ' + reqId + ' по подразделению «' + dept + '» на ' + year + ' год создана.\n\n';
  } else if (accepted.length) {
    text += 'Заявка ' + reqId + ' дополнена.\n\n';
  }
  if (accepted.length) {
    text += 'Принято позиций: ' + accepted.length + '\n';
    accepted.forEach(function (it, i) {
      text += (i + 1) + '. ' + it.name + (it.inCatalog ? '' : ' ⚠️ нет в каталоге') + ' — ' + it.qty + ' ' + it.unit + '\n';
    });
  }
  if (rejected.length) {
    text += '\nНе приняты (исправьте и отправьте повторно):\n';
    rejected.forEach(function (r) {
      text += '• ' + r.name + ' — ' + r.reason + '\n';
    });
  }
  if (!accepted.length && !rejected.length) {
    text = 'Заявка получена, но не содержит позиций.';
  }

  sendMessage(chatId, text);
}

// ==================== РАБОТА С ЛИСТАМИ ====================

function getCatalog() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATALOG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (v) { return v; });
}

function findActiveRequest(dept, year) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REQUESTS);
  if (sheet.getLastRow() < 2) return null;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[1] === dept && Number(row[2]) === Number(year) && row[3] !== 'Закрыта') {
      return { id: row[0], count: row[6], updated: row[5] };
    }
  }
  return null;
}

function generateRequestId(deptCode, year) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REQUESTS);
  var count = sheet.getLastRow() < 2 ? 0 : sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .getValues()
    .filter(function (r) { return String(r[0]).indexOf(deptCode + '-' + year) === 0; })
    .length;
  var seq = String(count + 1).padStart(3, '0');
  return deptCode + '-' + year + '-' + seq;
}

function addRequestRow(reqId, dept, year, chatId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REQUESTS);
  var now = new Date();
  sheet.appendRow([reqId, dept, year, 'Черновик', now, now, 0, chatId]);
}

function appendItems(reqId, items) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ITEMS);
  var startNo = getItemsForRequest(reqId).length + 1;
  var rows = items.map(function (it, i) {
    return [reqId, startNo + i, it.name, it.inCatalog ? 'Да' : 'Нет', it.justification, it.qty, it.unit, it.link];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
}

function getItemsForRequest(reqId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ITEMS);
  if (sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  return data.filter(function (r) { return r[0] === reqId; })
    .map(function (r) {
      return { name: r[2], inCatalog: r[3] === 'Да', justification: r[4], qty: r[5], unit: r[6], link: r[7] };
    });
}

function updateRequestSummary(reqId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REQUESTS);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === reqId) {
      var rowIndex = i + 2;
      var count = getItemsForRequest(reqId).length;
      sheet.getRange(rowIndex, 6).setValue(new Date()); // Изменена
      sheet.getRange(rowIndex, 7).setValue(count);       // Кол-во позиций
      if (count >= MAX_ITEMS_PER_REQUEST) {
        sheet.getRange(rowIndex, 4).setValue('Черновик (лимит)');
      }
      break;
    }
  }
}

// ==================== TELEGRAM API ====================

function sendMessage(chatId, text, replyMarkup) {
  var payload = { chat_id: chatId, text: text };
  if (replyMarkup) payload.reply_markup = JSON.stringify(replyMarkup);
  UrlFetchApp.fetch('https://api.telegram.org/bot' + getBotToken() + '/sendMessage', {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });
}

function answerCallbackQuery(id) {
  UrlFetchApp.fetch('https://api.telegram.org/bot' + getBotToken() + '/answerCallbackQuery', {
    method: 'post',
    payload: { callback_query_id: id },
    muteHttpExceptions: true
  });
}
