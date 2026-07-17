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
  for (
    let monthIndex = 0;
    monthIndex < LAYOUT.MONTH.COUNT;
    monthIndex++
  ) {
    const month = getMonthLayoutV1_(monthIndex);

    sheet.getRange(month.kpiRow, 1, LAYOUT.QUARTER.KPI_COUNT, 1).clearContent();
    sheet.getRange(month.kpiRow, 33, LAYOUT.QUARTER.KPI_COUNT, 1).clearContent();
    sheet.getRange(month.kpiRow, 47, LAYOUT.QUARTER.KPI_COUNT, 1).clearContent();
    sheet.getRange(month.kpiRow, 61, LAYOUT.QUARTER.KPI_COUNT, 1).clearContent();

    for (
      let sprintIndex = 0;
      sprintIndex < LAYOUT.SPRINT.COUNT;
      sprintIndex++
    ) {
      const sprint = getSprintLayoutV1_(
        monthIndex,
        sprintIndex
      );
      const firstColumn = LAYOUT.HYPOTHESIS.COLUMNS.NAME;
      const columnCount =
        LAYOUT.HYPOTHESIS.COLUMNS.RESULT - firstColumn + 1;

      sheet
        .getRange(
          sprint.firstHypRow,
          firstColumn,
          LAYOUT.SPRINT.HYPOTHESIS_COUNT,
          columnCount
        )
        .clearContent();
    }
  }

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

  sheet.getRange(LAYOUT.PASSPORT.ACCOUNT).setValue(data.account ?? "");
  sheet.getRange(LAYOUT.PASSPORT.SOCIAL).setValue(data.social ?? "");
  sheet.getRange(LAYOUT.PASSPORT.MANAGER).setValue(data.manager ?? "");
  if (String(data.status ?? '').trim() !== '') {
    sheet.getRange(LAYOUT.PASSPORT.STATUS).setValue(data.status);
  }

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
  return createAccountSheetsBatchV1();
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
  const platforms = getPlatformAbbreviationLookupV1_();
  const platform = platforms[normalizeDirectoryKeyV1_(social)];

  if (!platform) {
    throw new Error(`Для соцсети "${String(social).trim()}" не найдено сокращение.`);
  }

  return buildAccountSheetNameV1_(id, platform, account);
}

function normalizeDirectoryKeyV1_(value) {
  return String(value ?? '').trim().toLocaleLowerCase();
}

function getPlatformAbbreviationLookupV1_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DIRECTORY);

  if (!sheet) {
    throw new Error(`Лист "${SHEETS.DIRECTORY}" не найден.`);
  }

  const values = sheet.getDataRange().getDisplayValues();
  let headerRow = -1;
  let socialColumn = -1;
  let abbreviationColumn = -1;

  for (let row = 0; row < values.length && headerRow < 0; row++) {
    const headers = values[row].map(normalizeDirectoryKeyV1_);
    const socialIndex = headers.indexOf('соцсеть');
    const abbreviationIndex = headers.indexOf('сокращение');

    if (socialIndex >= 0 && abbreviationIndex >= 0) {
      headerRow = row;
      socialColumn = socialIndex;
      abbreviationColumn = abbreviationIndex;
    }
  }

  if (headerRow < 0) {
    throw new Error(
      `На листе "${SHEETS.DIRECTORY}" не найдены заголовки "Соцсеть" и "Сокращение".`
    );
  }

  const lookup = {};

  for (let row = headerRow + 1; row < values.length; row++) {
    const social = String(values[row][socialColumn] ?? '').trim();
    const abbreviation = String(values[row][abbreviationColumn] ?? '').trim();

    if (!social || !abbreviation) continue;
    lookup[normalizeDirectoryKeyV1_(social)] = abbreviation;
  }

  return lookup;
}

function normalizeAccountComponentV1_(account) {
  const normalized = String(account ?? '')
    .trim()
    .toLocaleLowerCase()
    .replace(/\s+/g, '_')
    .replace(/[:\\/?*\[\]]/g, '_')
    .replace(/[\u0000-\u001f\u007f]/g, '')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');

  if (!normalized) {
    throw new Error('После нормализации имя аккаунта оказалось пустым.');
  }

  return normalized;
}

function buildAccountSheetNameV1_(id, platform, account) {
  const rawId = String(id ?? '').trim();
  const safePlatform = String(platform ?? '').trim();

  if (!rawId) throw new Error('Не указан ID аккаунта.');
  if (!safePlatform) throw new Error('Не указано сокращение соцсети.');

  const safeId = rawId.padStart(3, '0');
  const prefix = `${SHEETS.ACCOUNT_PREFIX}${safeId}_${safePlatform}_`;
  const maxAccountLength =
    ACCOUNT_GENERATOR.MAX_SHEET_NAME_LENGTH - prefix.length;

  if (maxAccountLength < 1) {
    throw new Error('Префикс имени листа превышает допустимую длину.');
  }

  const safeAccount = normalizeAccountComponentV1_(account)
    .substring(0, maxAccountLength)
    .replace(/_+$/g, '');

  if (!safeAccount) {
    throw new Error('Не удалось сформировать компонент аккаунта для имени листа.');
  }

  return `${prefix}${safeAccount}`;
}

function getAccountRegistrySheetV1_() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEETS.ACCOUNTS);

  if (!sheet) throw new Error(`Лист "${SHEETS.ACCOUNTS}" не найден.`);
  return sheet;
}

function getAccountRegistryRowDataV1_(rowNumber, values) {
  const columns = ACCOUNT_REGISTRY.COLUMNS;
  const value = column => values[column - 1];

  return {
    rowNumber,
    values,
    id: value(columns.ID),
    sheetId: value(columns.SHEET_ID),
    sheetName: String(value(columns.SHEET_NAME) ?? '').trim(),
    social: String(value(columns.SOCIAL) ?? '').trim(),
    account: String(value(columns.ACCOUNT) ?? '').trim(),
    manager: String(value(columns.MANAGER) ?? '').trim(),
    year: value(columns.YEAR),
    quarter: value(columns.QUARTER),
    status: value(columns.STATUS),
    url: String(value(columns.URL) ?? '').trim(),
    createdAt: value(columns.CREATED_AT),
  };
}

function isAccountRegistryRowBlankV1_(rowData) {
  return rowData.values.every(value => String(value ?? '').trim() === '');
}

function validateAccountRegistryRowV1_(rowData, platformLookup) {
  const missing = [];
  const required = [
    ['ID', rowData.id],
    ['Соцсеть', rowData.social],
    ['Аккаунт', rowData.account],
    ['Ответственный', rowData.manager],
  ];

  required.forEach(([label, value]) => {
    if (String(value ?? '').trim() === '') missing.push(label);
  });

  if (missing.length) {
    throw new Error(`Не заполнены обязательные поля: ${missing.join(', ')}.`);
  }

  const platform = platformLookup[normalizeDirectoryKeyV1_(rowData.social)];
  if (!platform) {
    throw new Error(`Для соцсети "${rowData.social}" не найдено сокращение.`);
  }

  const expectedName = buildAccountSheetNameV1_(
    rowData.id,
    platform,
    rowData.account
  );

  if (
    expectedName.length > ACCOUNT_GENERATOR.MAX_SHEET_NAME_LENGTH ||
    /[:\\/?*\[\]]/.test(expectedName)
  ) {
    throw new Error(`Сформировано недопустимое имя листа "${expectedName}".`);
  }

  return { platform, expectedName };
}

function findSheetByIdV1_(ss, sheetId) {
  const normalizedId = String(sheetId ?? '').trim();
  if (!normalizedId) return null;

  return ss.getSheets().find(
    sheet => String(sheet.getSheetId()) === normalizedId
  ) || null;
}

function getAccountSheetUrlV1_(ss, sheet) {
  return `${ss.getUrl()}#gid=${sheet.getSheetId()}`;
}

function isProtectedAccountDistributionSheetV1_(sheet) {
  const protectedNames = new Set([
    SHEETS.TEMPLATE,
    SHEETS.ACCOUNTS,
    SHEETS.DIRECTORY,
  ]);
  return !sheet ||
    protectedNames.has(sheet.getName()) ||
    !sheet.getName().startsWith(SHEETS.ACCOUNT_PREFIX);
}

function resolveRegistryAccountSheetV1_(ss, rowData) {
  const byId = findSheetByIdV1_(ss, rowData.sheetId);
  let protectedReference = false;
  if (byId && !isProtectedAccountDistributionSheetV1_(byId)) {
    const expectedUrl = getAccountSheetUrlV1_(ss, byId);
    return {
      sheet: byId,
      reconciled: rowData.url !== expectedUrl,
    };
  }
  if (byId) protectedReference = true;

  const byName = rowData.sheetName
    ? ss.getSheetByName(rowData.sheetName)
    : null;
  if (byName && !isProtectedAccountDistributionSheetV1_(byName)) {
    return {
      sheet: byName,
      reconciled:
        String(rowData.sheetId ?? '').trim() !== String(byName.getSheetId()) ||
        rowData.url !== getAccountSheetUrlV1_(ss, byName),
    };
  }
  if (byName) protectedReference = true;

  if (protectedReference) return { sheet: null, skipped: true };

  throw new Error(
    `Лист аккаунта не найден по SheetId "${rowData.sheetId || ''}" ` +
    `или имени "${rowData.sheetName || ''}".`
  );
}

function persistKpiDistributionResultV1_(registry, rowNumber, writes) {
  const columns = ACCOUNT_REGISTRY.COLUMNS;
  if (writes.sheet) {
    registry.getRange(rowNumber, columns.SHEET_ID)
      .setValue(writes.sheet.getSheetId());
    registry.getRange(rowNumber, columns.URL)
      .setValue(getAccountSheetUrlV1_(SpreadsheetApp.getActiveSpreadsheet(), writes.sheet));
  }
  registry.getRange(rowNumber, columns.UPDATED_AT).setValue(new Date());
  registry.getRange(rowNumber, columns.ERROR).setValue(writes.error || '');
  SpreadsheetApp.flush();
}

/**
 * Distributes current quarterly KPI plans across months for registry accounts.
 * Progress is resumable through Script Properties.
 */
function distributeKpiPlansForAllAccountsBatchV1() {
  requireAdmin_();

  return withAccountGeneratorLockV1_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const registry = getAccountRegistrySheetV1_();
    const properties = PropertiesService.getScriptProperties();
    const cursorKey = KPI_DISTRIBUTION_BATCH.CURSOR_PROPERTY;
    const firstRow = ACCOUNT_REGISTRY.FIRST_DATA_ROW;
    const lastRow = registry.getLastRow();
    const storedRow = Number(properties.getProperty(cursorKey));
    const startRow = Number.isInteger(storedRow) &&
      storedRow >= firstRow && storedRow <= lastRow
      ? storedRow
      : firstRow;
    const startedAt = Date.now();
    const summary = {
      startRow,
      lastProcessedRow: null,
      success: 0,
      reconciled: 0,
      errors: 0,
      skippedBlank: 0,
      processed: 0,
      complete: lastRow < firstRow,
      nextRow: null,
    };

    for (let rowNumber = startRow; rowNumber <= lastRow; rowNumber++) {
      if (
        Date.now() - startedAt >= KPI_DISTRIBUTION_BATCH.MAX_RUNTIME_MS ||
        summary.processed >= KPI_DISTRIBUTION_BATCH.MAX_PROCESSED_PER_RUN
      ) {
        summary.nextRow = rowNumber;
        break;
      }

      const values = registry
        .getRange(rowNumber, 1, 1, ACCOUNT_REGISTRY.COLUMN_COUNT)
        .getValues()[0];
      const rowData = getAccountRegistryRowDataV1_(rowNumber, values);
      summary.lastProcessedRow = rowNumber;

      if (isAccountRegistryRowBlankV1_(rowData)) {
        summary.skippedBlank++;
        properties.setProperty(cursorKey, String(rowNumber + 1));
        continue;
      }

      let countedAsProcessed = false;
      try {
        const resolved = resolveRegistryAccountSheetV1_(ss, rowData);
        if (resolved.skipped) {
          properties.setProperty(cursorKey, String(rowNumber + 1));
          continue;
        }
        summary.processed++;
        countedAsProcessed = true;
        if (resolved.reconciled) {
          persistKpiDistributionResultV1_(registry, rowNumber, {
            sheet: resolved.sheet,
            error: '',
          });
          summary.reconciled++;
        }
        distributeKpiPlansForSheetV1_(resolved.sheet);
        persistKpiDistributionResultV1_(registry, rowNumber, {
          sheet: null,
          error: '',
        });
        summary.success++;
      } catch (error) {
        if (!countedAsProcessed) summary.processed++;
        persistKpiDistributionResultV1_(registry, rowNumber, {
          error: String(error && error.message ? error.message : error),
        });
        summary.errors++;
      }

      properties.setProperty(cursorKey, String(rowNumber + 1));
    }

    if (!summary.nextRow) {
      summary.complete = true;
      properties.deleteProperty(cursorKey);
    } else {
      properties.setProperty(cursorKey, String(summary.nextRow));
    }

    SpreadsheetApp.getUi().alert(
      'Распределение KPI завершено.\n\n' +
      `Стартовая строка: ${summary.startRow}\n` +
      `Последняя обработанная: ${summary.lastProcessedRow || 'нет'}\n` +
      `Успешно: ${summary.success}\n` +
      `Сверено ссылок: ${summary.reconciled}\n` +
      `Ошибок: ${summary.errors}\n` +
      `Пустых строк пропущено: ${summary.skippedBlank}\n` +
      `Полный проход завершён: ${summary.complete ? 'да' : 'нет'}\n` +
      `Следующая строка: ${summary.nextRow || 'сброшено'}`
    );

    return summary;
  });
}

function finalizeAccountSheetV1_(sheet, sheetName, rowData) {
  if (sheet.getName() !== sheetName) {
    throw new Error(
      `Имя найденного листа "${sheet.getName()}" не совпадает с ожидаемым "${sheetName}".`
    );
  }

  sheet.getRange(1, 1).setValue(sheetName);
  fillPassportV1_(sheet, rowData);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(0);
}

function persistAccountRegistrySuccessV1_(registry, rowData, sheet) {
  const columns = ACCOUNT_REGISTRY.COLUMNS;
  const now = new Date();
  const createdAt = rowData.createdAt || now;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateYear = sheet
    .getRange(LAYOUT.PASSPORT.YEAR)
    .getValue();
  const templateQuarter = sheet
    .getRange(LAYOUT.PASSPORT.QUARTER)
    .getValue();
  const writes = [
    [columns.SHEET_ID, sheet.getSheetId()],
    [columns.SHEET_NAME, sheet.getName()],
    [columns.YEAR, templateYear],
    [columns.QUARTER, templateQuarter],
    [columns.STATUS, rowData.status],
    [columns.URL, getAccountSheetUrlV1_(ss, sheet)],
    [columns.CREATED_AT, createdAt],
    [columns.UPDATED_AT, now],
    [columns.ERROR, ''],
  ];

  let lastError = null;
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      writes.forEach(([column, value]) => {
        registry.getRange(rowData.rowNumber, column).setValue(value);
      });
      SpreadsheetApp.flush();
      return;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError;
}

function persistAccountRegistryErrorV1_(registry, rowNumber, error) {
  const columns = ACCOUNT_REGISTRY.COLUMNS;
  const message = String(error && error.message ? error.message : error)
    .trim()
    .substring(0, 500);

  registry
    .getRange(rowNumber, columns.UPDATED_AT)
    .setValue(new Date());
  registry
    .getRange(rowNumber, columns.ERROR)
    .setValue(message || 'Неизвестная ошибка создания листа.');
  SpreadsheetApp.flush();
}

function createAccountSheetCopyV1_(ss, template, expectedName, rowData) {
  let sheet = null;
  let renamed = false;

  try {
    sheet = template.copyTo(ss);
    sheet.setName(expectedName);
    renamed = true;
    finalizeAccountSheetV1_(sheet, expectedName, rowData);
    SpreadsheetApp.flush();
    return sheet;
  } catch (error) {
    if (sheet && !renamed) {
      try {
        ss.deleteSheet(sheet);
      } catch (deleteError) {}
    }
    throw error;
  }
}

function processAccountRegistryRowV1_(context, rowNumber, values) {
  const { ss, registry, template, platformLookup } = context;
  const rowData = getAccountRegistryRowDataV1_(rowNumber, values);
  const { expectedName } = validateAccountRegistryRowV1_(
    rowData,
    platformLookup
  );
  const storedSheet = findSheetByIdV1_(ss, rowData.sheetId);

  if (storedSheet) {
    const expectedUrl = getAccountSheetUrlV1_(ss, storedSheet);
    const alreadyComplete =
      storedSheet.getName() === expectedName &&
      rowData.sheetName === expectedName &&
      rowData.url === expectedUrl;

    if (!alreadyComplete) {
      if (storedSheet.getName() !== expectedName) {
        throw new Error(
          `SheetId ${rowData.sheetId} указывает на лист "${storedSheet.getName()}", ожидался "${expectedName}".`
        );
      }
      finalizeAccountSheetV1_(storedSheet, expectedName, rowData);
      persistAccountRegistrySuccessV1_(registry, rowData, storedSheet);
      return { status: 'reconciled', sheet: storedSheet, expectedName };
    }

    return { status: 'skipped', sheet: storedSheet, expectedName };
  }

  const sheetByName = ss.getSheetByName(expectedName);

  if (sheetByName) {
    finalizeAccountSheetV1_(sheetByName, expectedName, rowData);
    persistAccountRegistrySuccessV1_(registry, rowData, sheetByName);
    return { status: 'reconciled', sheet: sheetByName, expectedName };
  }

  if (String(rowData.sheetId ?? '').trim()) {
    throw new Error(
      `SheetId ${rowData.sheetId} не найден; лист "${expectedName}" также отсутствует.`
    );
  }

  const sheet = createAccountSheetCopyV1_(
    ss,
    template,
    expectedName,
    rowData
  );

  persistAccountRegistrySuccessV1_(registry, rowData, sheet);
  return { status: 'created', sheet, expectedName };
}

function getAccountGeneratorContextV1_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registry = getAccountRegistrySheetV1_();
  const template = ss.getSheetByName(SHEETS.TEMPLATE);

  if (!template) throw new Error(`Лист "${SHEETS.TEMPLATE}" не найден.`);

  return {
    ss,
    registry,
    template,
    platformLookup: getPlatformAbbreviationLookupV1_(),
  };
}

function withAccountGeneratorLockV1_(callback) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    throw new Error('Генератор листов уже выполняется другим пользователем.');
  }

  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function createNextAccountSheetTestV1() {
  return withAccountGeneratorLockV1_(() => {
    const context = getAccountGeneratorContextV1_();
    const { registry } = context;
    const lastRow = registry.getLastRow();

    if (lastRow < ACCOUNT_REGISTRY.FIRST_DATA_ROW) {
      SpreadsheetApp.getUi().alert('В реестре нет строк аккаунтов.');
      return;
    }

    const values = registry
      .getRange(
        ACCOUNT_REGISTRY.FIRST_DATA_ROW,
        1,
        lastRow - ACCOUNT_REGISTRY.FIRST_DATA_ROW + 1,
        ACCOUNT_REGISTRY.COLUMN_COUNT
      )
      .getValues();

    for (let index = 0; index < values.length; index++) {
      const rowNumber = ACCOUNT_REGISTRY.FIRST_DATA_ROW + index;
      const rowData = getAccountRegistryRowDataV1_(rowNumber, values[index]);
      if (isAccountRegistryRowBlankV1_(rowData)) continue;

      const storedSheet = findSheetByIdV1_(context.ss, rowData.sheetId);
      if (storedSheet) {
        try {
          const { expectedName } = validateAccountRegistryRowV1_(
            rowData,
            context.platformLookup
          );
          const expectedUrl = getAccountSheetUrlV1_(context.ss, storedSheet);

          if (
            storedSheet.getName() === expectedName &&
            rowData.sheetName === expectedName &&
            rowData.url === expectedUrl
          ) {
            continue;
          }
        } catch (error) {}
      }

      let expectedName = '';

      try {
        expectedName = validateAccountRegistryRowV1_(
          rowData,
          context.platformLookup
        ).expectedName;
        const result = processAccountRegistryRowV1_(
          context,
          rowNumber,
          values[index]
        );
        const url = getAccountSheetUrlV1_(context.ss, result.sheet);

        SpreadsheetApp.getUi().alert(
          `Успешно (${result.status}).\n\n` +
          `Строка: ${rowNumber}\n` +
          `Лист: ${result.sheet.getName()}\n` +
          `SheetId: ${result.sheet.getSheetId()}\n` +
          `URL: ${url}`
        );
      } catch (error) {
        persistAccountRegistryErrorV1_(registry, rowNumber, error);
        const existingSheet = expectedName
          ? context.ss.getSheetByName(expectedName)
          : null;
        SpreadsheetApp.getUi().alert(
          `Ошибка.\n\n` +
          `Строка: ${rowNumber}\n` +
          `Лист: ${expectedName || 'не сформирован'}\n` +
          `SheetId: ${existingSheet ? existingSheet.getSheetId() : 'нет'}\n` +
          `URL: ${existingSheet ? getAccountSheetUrlV1_(context.ss, existingSheet) : 'нет'}\n` +
          `${error.message}`
        );
      }
      return;
    }

    SpreadsheetApp.getUi().alert(
      'Не найдено строк, ожидающих создания или сверки.'
    );
  });
}

function createAccountSheetsBatchV1(maxCreatedPerRun) {
  return withAccountGeneratorLockV1_(() => {
    const context = getAccountGeneratorContextV1_();
    const { registry } = context;
    const startedAt = Date.now();
    const maxRuntimeMs = ACCOUNT_GENERATOR.MAX_RUNTIME_MS;
    const createLimit = Number(maxCreatedPerRun) > 0
      ? Math.floor(Number(maxCreatedPerRun))
      : ACCOUNT_GENERATOR.MAX_CREATED_PER_RUN;
    const lastRow = registry.getLastRow();
    const summary = {
      created: 0,
      reconciled: 0,
      skipped: 0,
      errors: 0,
      nextRow: null,
    };

    if (lastRow >= ACCOUNT_REGISTRY.FIRST_DATA_ROW) {
      const values = registry
        .getRange(
          ACCOUNT_REGISTRY.FIRST_DATA_ROW,
          1,
          lastRow - ACCOUNT_REGISTRY.FIRST_DATA_ROW + 1,
          ACCOUNT_REGISTRY.COLUMN_COUNT
        )
        .getValues();

      for (let index = 0; index < values.length; index++) {
        const rowNumber = ACCOUNT_REGISTRY.FIRST_DATA_ROW + index;

        if (
          Date.now() - startedAt >= maxRuntimeMs ||
          summary.created >= createLimit
        ) {
          summary.nextRow = rowNumber;
          break;
        }

        const rowData = getAccountRegistryRowDataV1_(rowNumber, values[index]);
        if (isAccountRegistryRowBlankV1_(rowData)) {
          summary.skipped++;
          continue;
        }

        try {
          const result = processAccountRegistryRowV1_(
            context,
            rowNumber,
            values[index]
          );
          summary[result.status]++;
        } catch (error) {
          summary.errors++;
          if (summary.nextRow === null) summary.nextRow = rowNumber;
          try {
            persistAccountRegistryErrorV1_(registry, rowNumber, error);
          } catch (registryError) {}
        }
      }
    }

    SpreadsheetApp.getUi().alert(
      'Пакетное создание завершено.\n\n' +
      `Создано: ${summary.created}\n` +
      `Сверено: ${summary.reconciled}\n` +
      `Пропущено: ${summary.skipped}\n` +
      `Ошибок: ${summary.errors}\n` +
      `Следующая строка: ${summary.nextRow || 'нет'}`
    );

    return summary;
  });
}

function findAccountRegistryRowForSheetV1_(registryValues, sheetId, sheetName) {
  const columns = ACCOUNT_REGISTRY.COLUMNS;
  const idIndex = columns.SHEET_ID - 1;
  const nameIndex = columns.SHEET_NAME - 1;
  const normalizedId = String(sheetId ?? '').trim();

  let index = registryValues.findIndex(row =>
    String(row[idIndex] ?? '').trim() === normalizedId
  );

  if (index < 0) {
    index = registryValues.findIndex(row =>
      String(row[nameIndex] ?? '').trim() === sheetName
    );
  }

  return index < 0 ? null : ACCOUNT_REGISTRY.FIRST_DATA_ROW + index;
}

function clearDeletedAccountRegistryLinkV1_(registry, rowNumber) {
  const columns = ACCOUNT_REGISTRY.COLUMNS;
  const clearColumns = [
    columns.SHEET_ID,
    columns.SHEET_NAME,
    columns.URL,
    columns.CREATED_AT,
    columns.ERROR,
  ];

  clearColumns.forEach(column => {
    registry.getRange(rowNumber, column).clearContent();
  });
  registry
    .getRange(rowNumber, columns.UPDATED_AT)
    .setValue(new Date());
  SpreadsheetApp.flush();
}

function writeAccountDeletionErrorV1_(registry, rowNumber, error) {
  if (!rowNumber) return;

  const columns = ACCOUNT_REGISTRY.COLUMNS;
  const message = String(error && error.message ? error.message : error)
    .trim()
    .substring(0, 500);

  registry
    .getRange(rowNumber, columns.UPDATED_AT)
    .setValue(new Date());
  registry
    .getRange(rowNumber, columns.ERROR)
    .setValue(message || 'Ошибка удаления листа аккаунта.');
  SpreadsheetApp.flush();
}

function deleteAllAccountSheetsV1() {
  requireAdmin_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const registry = getAccountRegistrySheetV1_();
  const protectedNames = new Set(
    Object.keys(SHEETS)
      .filter(key => key !== 'ACCOUNT_PREFIX')
      .map(key => SHEETS[key])
  );
  const sheetSnapshot = ss.getSheets().slice();
  const targets = sheetSnapshot.filter(sheet => {
    const name = sheet.getName();
    return name.startsWith(SHEETS.ACCOUNT_PREFIX) &&
      !protectedNames.has(name);
  });
  const skippedProtected = sheetSnapshot.filter(sheet =>
    protectedNames.has(sheet.getName())
  ).length;

  if (!targets.length) {
    ui.alert('Листы аккаунтов для удаления не найдены.');
    return;
  }

  const firstConfirmation = ui.alert(
    'Удаление листов аккаунтов',
    `Будут удалены все созданные листы аккаунтов (${targets.length}). ` +
      'Данные в этих листах будут потеряны. Продолжить?',
    ui.ButtonSet.YES_NO
  );

  if (firstConfirmation !== ui.Button.YES) {
    ui.alert('Удаление отменено.');
    return;
  }

  const confirmationPhrase = 'УДАЛИТЬ ВСЕ ЛИСТЫ';
  const secondConfirmation = ui.prompt(
    'Повторное подтверждение',
    `Введите точную фразу: ${confirmationPhrase}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (
    secondConfirmation.getSelectedButton() !== ui.Button.OK ||
    secondConfirmation.getResponseText() !== confirmationPhrase
  ) {
    ui.alert('Фраза подтверждения не совпала. Удаление отменено.');
    return;
  }

  const safeSheet = ss.getSheetByName(SHEETS.ACCOUNTS) ||
    ss.getSheetByName(SHEETS.TEMPLATE) ||
    ss.getSheetByName(SHEETS.DIRECTORY);

  if (!safeSheet) {
    throw new Error('Не найден защищённый лист для безопасной активации.');
  }

  safeSheet.activate();

  const lastRow = registry.getLastRow();
  const registryValues = lastRow >= ACCOUNT_REGISTRY.FIRST_DATA_ROW
    ? registry
        .getRange(
          ACCOUNT_REGISTRY.FIRST_DATA_ROW,
          1,
          lastRow - ACCOUNT_REGISTRY.FIRST_DATA_ROW + 1,
          ACCOUNT_REGISTRY.COLUMN_COUNT
        )
        .getValues()
    : [];
  const summary = {
    found: targets.length,
    deleted: 0,
    registryCleared: 0,
    errors: 0,
    skippedProtected,
  };

  targets.forEach(sheet => {
    const sheetId = sheet.getSheetId();
    const sheetName = sheet.getName();
    const registryRow = findAccountRegistryRowForSheetV1_(
      registryValues,
      sheetId,
      sheetName
    );

    try {
      ss.deleteSheet(sheet);
      summary.deleted++;

      if (registryRow) {
        try {
          clearDeletedAccountRegistryLinkV1_(registry, registryRow);
          summary.registryCleared++;
        } catch (registryError) {
          summary.errors++;
          try {
            writeAccountDeletionErrorV1_(
              registry,
              registryRow,
              registryError
            );
          } catch (writeError) {}
        }
      } else {
        summary.errors++;
      }
    } catch (deleteError) {
      summary.errors++;
      try {
        writeAccountDeletionErrorV1_(registry, registryRow, deleteError);
      } catch (writeError) {}
    }
  });

  ui.alert(
    'Удаление завершено.\n\n' +
    `Найдено листов: ${summary.found}\n` +
    `Удалено листов: ${summary.deleted}\n` +
    `Очищено строк реестра: ${summary.registryCleared}\n` +
    `Ошибок: ${summary.errors}\n` +
    `Пропущено защищённых листов: ${summary.skippedProtected}`
  );

  return summary;
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

  const names = reg.getRange(2, 3, lastRow - 1, 1).getValues();

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
