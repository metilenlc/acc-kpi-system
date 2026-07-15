/*
createAccountSheetV1()
getNextAccountIdV1_()
normalizeAccountNameV1_()
isAccountSheetV1_()
*/

/**
 * ==========================================
 * clearAccountDataV1_
 * ==========================================
 *
 * Очищает рабочие данные листа аккаунта,
 * не затрагивая структуру, формулы,
 * форматирование и выпадающие списки.
 */
function clearAccountDataV1_(sheet) {
  if (!sheet) return;

  // Паспорт
  sheet.getRange("M4:M10").clearContent();

  // KPI квартала: KPI, План, Факт, %
  sheet.getRange("A21:A32").clearContent();
  sheet.getRange("AG21:AG32").clearContent();
  sheet.getRange("AU21:AU32").clearContent();
  sheet.getRange("BI21:BI32").clearContent();

  // KPI месяцев: KPI, План, Факт, %
  const monthKpiRows = [48, 156, 264];

  monthKpiRows.forEach(row => {
    sheet.getRange(row, 1, 12, 1).clearContent();   // KPI
    sheet.getRange(row, 33, 12, 1).clearContent();  // План
    sheet.getRange(row, 47, 12, 1).clearContent();  // Факт
    sheet.getRange(row, 61, 12, 1).clearContent();  // %
  });

  // Гипотезы
  const monthHypRows = [78, 186, 294];

  monthHypRows.forEach(monthStart => {
    for (let sprint = 0; sprint < 5; sprint++) {
      const row = monthStart + sprint * 15;
      sheet.getRange(row, 3, 4, 81).clearContent();
    }
  });

  SpreadsheetApp.flush();
}

/**
 * ==========================================
 * copyTemplateSheetV1_
 * ==========================================
 *
 * Создаёт копию листа ACC_TEMPLATE.
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function copyTemplateSheetV1_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const template = ss.getSheetByName("ACC_TEMPLATE");

  if (!template) {
    throw new Error('Лист "ACC_TEMPLATE" не найден.');
  }

  const sheet = template.copyTo(ss);

  SpreadsheetApp.flush();

  return sheet;

}

/**
 * ==========================================
 * renameAccountSheetV1_
 * ==========================================
 *
 * Переименовывает лист аккаунта.
 *
 * @param {Sheet} sheet
 * @param {string} accountName
 * @return {Sheet}
 */
function renameAccountSheetV1_(sheet, accountName) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  accountName = String(accountName).trim();

  if (!accountName) {
    throw new Error("Не указано имя аккаунта.");
  }

  if (ss.getSheetByName(accountName)) {
    throw new Error(`Лист "${accountName}" уже существует.`);
  }

  sheet.setName(accountName);

  SpreadsheetApp.flush();

  return sheet;

}

/**
 * ==========================================
 * fillPassportV1_
 * ==========================================
 *
 * Заполняет паспорт аккаунта.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} data
 */
function fillPassportV1_(sheet, data) {

  if (!sheet) throw new Error("Лист не передан.");

  data = data || {};

  const values = [
    [data.account ?? ""],   // M4
    [data.social ?? ""],    // M5
    [data.manager ?? ""],   // M6
    [data.year ?? ""],      // M7
    [data.quarter ?? ""],   // M8
    [data.status ?? ""]     // M9
  ];

  sheet.getRange(4, 13, values.length, 1).setValues(values);

}

/**
 * ==========================================
 * registerAccountV1_
 * ==========================================
 *
 * Регистрирует аккаунт в листе ACCOUNTS.
 *
 * Структура:
 * A - ID
 * B - SheetId
 * C - Лист
 * D - Аккаунт
 * E - Соцсеть
 * F - Ответственный
 * G - Статус
 * H - Создан
 * I - Обновлён
 */
/**
 * ==========================================
 * registerAccountV1_
 * ==========================================
 *
 * Регистрирует или обновляет аккаунт
 * в листе ACCOUNTS.
 */
function registerAccountV1_(sheet, data) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reg = ss.getSheetByName("ACCOUNTS");

  if (!reg) {
    throw new Error('Лист "ACCOUNTS" не найден.');
  }

  const lastRow = reg.getLastRow();

  if (lastRow < 2) {

    reg.appendRow([
      1,
      sheet.getSheetId(),
      sheet.getName(),
      data.account || "",
      data.social || "",
      data.manager || "",
      data.status || "",
      new Date(),
      new Date()
    ]);

    return;
  }

  const values = reg.getRange(2, 1, lastRow - 1, 9).getValues();

  for (let i = 0; i < values.length; i++) {

    if (values[i][3] === data.account) {

      values[i][1] = sheet.getSheetId();
      values[i][2] = sheet.getName();
      values[i][4] = data.social || "";
      values[i][5] = data.manager || "";
      values[i][6] = data.status || "";
      values[i][8] = new Date();

      reg.getRange(i + 2, 1, 1, 9).setValues([values[i]]);
      return;
    }
  }

  const id = values.length + 1;

  reg.getRange(lastRow + 1, 1, 1, 9).setValues([[
    id,
    sheet.getSheetId(),
    sheet.getName(),
    data.account || "",
    data.social || "",
    data.manager || "",
    data.status || "",
    new Date(),
    new Date()
  ]]);

}

/**
 * ==========================================
 * createAccountsV1
 * ==========================================
 *
 * Массовое создание аккаунтов
 * из листа ACCOUNTS.
 */
function createAccountV1(data) {
  if (!data || !data.account) {
    throw new Error("Не указано название аккаунта.");
  }

  let sheet = null;

  try {
    sheet = copyTemplateSheetV1_();

    const sheetName = buildSheetNameV1_(
      data.id,
      data.social,
      data.account
    );

  renameAccountSheetV1_(sheet, sheetName);

    sheet.getRange("A1").setValue(sheet.getName());

    fillPassportV1_(sheet, data);

    SpreadsheetApp.setActiveSheet(sheet);

    setupDropdownsV1();

    clearRowGroupsV1();
    setupQuarterKpiSummaryGroupV1();
    setupMonthGroupsV1();
    setupMonthKpiSummaryGroupsV1();
    setupSprintRowGroupsV1();

    collapseAllGroupsV1_(sheet);

    recalculateAnalyticsV1();

    registerAccountV1_(sheet, data);

    sheet.activate();

    return sheet;

  } catch (err) {
    if (sheet) {
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
    }

    throw err;
  }
}

function createAccountsV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reg = ss.getSheetByName("ACCOUNTS");

  if (!reg) throw new Error('Лист "ACCOUNTS" не найден.');

  const lastRow = reg.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Нет строк для создания аккаунтов.");
    return;
  }

  const values = reg.getRange(2, 1, lastRow - 1, 12).getValues();

  let created = 0;
  let skipped = 0;
  let errors = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const account = String(row[3] || "").trim(); // D
    const social = row[4];                       // E
    const manager = row[5];                      // F
    const year = row[6];                         // G
    const quarter = row[7];                      // H
    const status = row[8];                       // I

    if (!account) continue;

    if (row[1] || ss.getSheetByName(account)) {
      skipped++;
      row[11] = "";
      continue;
    }

    try {
      createAccountV1({
        id: row[0],
        account,
        social,
        manager,
        year,
        quarter,
        status
      });

      row[11] = "";
      created++;

    } catch (err) {
      row[11] = err.message;
      errors++;
    }
  }

  reg.getRange(2, 12, values.length, 1).setValues(
    values.map(r => [r[11]])
  );

  SpreadsheetApp.getUi().alert(
    "Создание аккаунтов завершено.\n\n" +
    "Создано: " + created + "\n" +
    "Пропущено: " + skipped + "\n" +
    "Ошибок: " + errors
  );
}

//группы начало
function collapseAllGroupsV1_(sheet) {
  if (!sheet) return;

  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
         .collapseGroups();
  } catch (e) {}
}

function expandAllGroupsV1_(sheet) {
  if (!sheet) return;

  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
         .expandGroups();
  } catch (e) {}
}

function collapseActiveSheetGroupsV1() {
  collapseAllGroupsV1_(SpreadsheetApp.getActiveSheet());
}

function expandActiveSheetGroupsV1() {
  expandAllGroupsV1_(SpreadsheetApp.getActiveSheet());
}
//группы конец

function buildSheetNameV1_(id, social, account) {
  const platformCodes = {
    "Instagram": "IG",
    "Telegram": "TG",
    "VK": "VK",
    "YouTube": "YT",
    "Дзен": "DZ",
    "TikTok": "TT",
    "Facebook": "FB",
    "Сайт": "WEB"
  };

  const platform = platformCodes[String(social).trim()] || "OTH";

  const safeId = String(id).padStart(3, "0");

  const safeAccount = String(account)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "_")
    .replace(/[^\wа-яё_-]/gi, "")
    .substring(0, 40);

  return `ACC_${safeId}_${platform}_${safeAccount}`;
}

//навигация начало
/**
 * Открывает лист аккаунта по выбранной строке в ACCOUNTS.
 * Техническое имя листа хранится в колонке C.
 */
function openSelectedAccountV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registry = ss.getActiveSheet();

  if (registry.getName() !== 'ACCOUNTS') {
    SpreadsheetApp.getUi().alert(
      'Сначала перейдите на лист ACCOUNTS и выберите строку аккаунта.'
    );
    return;
  }

  const row = registry.getActiveRange().getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert('Выберите строку аккаунта ниже заголовков.');
    return;
  }

  const sheetName = String(
    registry.getRange(row, 3).getDisplayValue()
  ).trim();

  if (!sheetName) {
    SpreadsheetApp.getUi().alert(
      'В колонке C выбранной строки не указано имя листа.'
    );
    return;
  }

  const accountSheet = ss.getSheetByName(sheetName);

  if (!accountSheet) {
    SpreadsheetApp.getUi().alert(
      `Лист "${sheetName}" не найден.`
    );
    return;
  }

  accountSheet.activate();
  accountSheet.setActiveSelection('A1');
}

/**
 * Находит текущий лист аккаунта в реестре ACCOUNTS.
 */
function findCurrentAccountV1() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const currentSheet = ss.getActiveSheet();

  if (currentSheet.getName() === "ACCOUNTS") {
    return;
  }

  const reg = ss.getSheetByName("ACCOUNTS");

  if (!reg) {
    SpreadsheetApp.getUi().alert('Лист "ACCOUNTS" не найден.');
    return;
  }

  const lastRow = reg.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Реестр пуст.");
    return;
  }

  const names = reg.getRange(2, 2, lastRow - 1, 1).getValues();

  for (let i = 0; i < names.length; i++) {

    if (names[i][0] === currentSheet.getName()) {

      reg.activate();
      reg.setActiveSelection(`A${i + 2}`);

      SpreadsheetApp.flush();

      return;
    }
  }

  SpreadsheetApp.getUi().alert("Текущий аккаунт не найден в реестре.");

}

/**
 * Безусловно открывает лист ACCOUNTS.
 */
function openAccountsRegistryV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ACCOUNTS');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Лист "ACCOUNTS" не найден.');
    return;
  }

  sheet.activate();
  sheet.setActiveSelection('A1');
}
//навигация конец

//тестовые функции
/**
 * Проверки
 */
function testCopyTemplate() {

  const sheet = copyTemplateSheetV1_();

  Logger.log(sheet.getName());

}

function testRenameSheet() {

  const sheet = copyTemplateSheetV1_();

  renameAccountSheetV1_(sheet, "TEST_ACCOUNT");

}

function testFillPassportV1() {

  const sheet = copyTemplateSheetV1_();

  renameAccountSheetV1_(sheet, "TEST_ACCOUNT");

  fillPassportV1_(sheet, {
    account: "ООО Ромашка",
    social: "Instagram",
    manager: "Иванов И.И.",
    year: 2026,
    quarter: "Q3",
    status: "Активный"
  });

}

function testRegisterAccountV1() {

  const sheet = copyTemplateSheetV1_();

  renameAccountSheetV1_(sheet, "TEST_ACCOUNT");

  fillPassportV1_(sheet, {
    account: "ООО Ромашка",
    social: "Instagram",
    manager: "Иванов И.И.",
    year: 2026,
    quarter: "Q3",
    status: "Активный"
  });

  registerAccountV1_(sheet, {
    account: "ООО Ромашка",
    social: "Instagram",
    manager: "Иванов И.И.",
    status: "Активный"
  });

}

function testCreateAccountV1() {
  createAccountV1({
    account: "TEST_ACCOUNT",
    social: "Instagram",
    manager: "Иванов И.И.",
    year: 2026,
    quarter: "Q3",
    status: "Активный"
  });
}



