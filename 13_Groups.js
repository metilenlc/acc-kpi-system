/*
==================================================
10_Groups.gs
Группировки строк
==================================================
*/

/**
 * ==========================================
 * setupSprintRowGroupsV1
 * ==========================================
 *
 * Назначение:
 * Создаёт внутренние группировки для спринтов.
 *
 * Логика:
 * • Заголовок СПРИНТ № остаётся видимым.
 * • Тело спринта группируется.
 * • Каждый спринт можно свернуть отдельно.
 */
function setupSprintRowGroupsV1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACC_TEMPLATE');

  const FIRST_ROW = 1;
  const LAST_ROW = sheet.getLastRow();

  const SPRINT_GROUP_START_OFFSET = 1;
  const SPRINT_GROUP_ROWS_COUNT = 13;

  //clearRowGroupsV1_();

  for (let row = FIRST_ROW; row <= LAST_ROW; row++) {
    const value = String(sheet.getRange(row, 1).getDisplayValue()).trim();

    if (!value.startsWith('СПРИНТ №')) continue;

    const groupStartRow = row + SPRINT_GROUP_START_OFFSET;

    sheet
      .getRange(groupStartRow, 1, SPRINT_GROUP_ROWS_COUNT, 1)
      .shiftRowGroupDepth(1);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Группировки спринтов созданы',
    'ACC',
    3
  );
}

/**
 * ==========================================
 * clearRowGroupsV1
 * ==========================================
 *
 * Назначение:
 * Удаляет все группировки строк на ACC_TEMPLATE.
 */
function clearRowGroupsV1() {
  clearRowGroupsV1_();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Группировки строк удалены',
    'ACC',
    3
  );
}

/**
 * ==========================================
 * clearRowGroupsV1_
 * ==========================================
 *
 * Назначение:
 * Внутренняя функция удаления группировок строк.
 */
function clearRowGroupsV1_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACC_TEMPLATE');

  const MAX_GROUP_DEPTH = 8;
  const LAST_ROW = sheet.getMaxRows();

  try {
    sheet.expandAllRowGroups();
  } catch (e) {}

  for (let depth = 0; depth < MAX_GROUP_DEPTH; depth++) {
    try {
      sheet
        .getRange(1, 1, LAST_ROW, 1)
        .shiftRowGroupDepth(-1);
    } catch (e) {
      break;
    }
  }
}

/**
 * ==========================================
 * setupMonthKpiSummaryGroupsV1
 * ==========================================
 *
 * Назначение:
 * Создаёт группировки для блока KPI месяца + Итоги месяца.
 *
 * Логика:
 * • Заголовок МЕСЯЦ N остаётся видимым.
 * • KPI месяца и Итоги месяца сворачиваются одной группой.
 * • Спринты группируются отдельно другой функцией.
 */
function setupMonthKpiSummaryGroupsV1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACC_TEMPLATE');

  const FIRST_ROW = 1;
  const LAST_ROW = sheet.getLastRow();

  const GROUP_START_OFFSET = 1;
  const GROUP_END_BEFORE_SPRINT = true;

  for (let row = FIRST_ROW; row <= LAST_ROW; row++) {
    const value = String(sheet.getRange(row, 1).getDisplayValue()).trim();

    if (!value.startsWith('МЕСЯЦ ')) continue;

    const monthTitleRow = row;
    const groupStartRow = monthTitleRow + GROUP_START_OFFSET;

    let firstSprintRow = null;

    for (let scanRow = groupStartRow; scanRow <= LAST_ROW; scanRow++) {
      const scanValue = String(sheet.getRange(scanRow, 1).getDisplayValue()).trim();

      if (scanValue.startsWith('СПРИНТ №')) {
        firstSprintRow = scanRow;
        break;
      }
    }

    if (!firstSprintRow) continue;

    const groupEndRow = firstSprintRow - 2;
    const groupRowsCount = groupEndRow - groupStartRow + 1;

    if (groupRowsCount <= 0) continue;

    sheet
      .getRange(groupStartRow, 1, groupRowsCount, 1)
      .shiftRowGroupDepth(1);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Группировки KPI и итогов месяца созданы',
    'ACC',
    3
  );
}

/**
 * ==========================================
 * setupMonthGroupsV1
 * ==========================================
 *
 * Назначение:
 * Создаёт внешнюю группировку всего месяца.
 *
 * Уровни:
 * 1 — месяц
 * 2 — KPI + Итоги
 * 3 — Спринты
 */
function setupMonthGroupsV1() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACC_TEMPLATE');

  const FIRST_ROW = 1;
  const LAST_ROW = sheet.getLastRow();

  for (let row = FIRST_ROW; row <= LAST_ROW; row++) {

    const value = String(sheet.getRange(row, 1).getDisplayValue()).trim();

    if (!value.startsWith('МЕСЯЦ ')) continue;

    const monthTitleRow = row;

    let nextMonthRow = LAST_ROW + 1;

    for (let r = monthTitleRow + 1; r <= LAST_ROW; r++) {

      const text = String(sheet.getRange(r, 1).getDisplayValue()).trim();

      if (text.startsWith('МЕСЯЦ ')) {
        nextMonthRow = r;
        break;
      }
    }

    const groupStart = monthTitleRow + 1;
    const groupRows = nextMonthRow - monthTitleRow - 2;

    if (groupRows <= 0) continue;

    sheet
      .getRange(groupStart, 1, groupRows, 1)
      .shiftRowGroupDepth(1);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Группировки месяцев созданы',
    'ACC',
    3
  );
}

/**
 * ==========================================
 * setupQuarterKpiSummaryGroupV1
 * ==========================================
 *
 * Назначение:
 * Создаёт группировку верхнего квартального блока:
 * • Квартальные KPI
 * • Итоги квартала
 *
 * Диапазон:
 * • строки 20–41
 * • зелёный заголовок "КВАРТАЛЬНЫЕ KPI" остаётся видимым
 */
function setupQuarterKpiSummaryGroupV1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACC_TEMPLATE');

  const GROUP_START_ROW = 20;
  const GROUP_ROWS_COUNT = 22; // 20–41

  sheet
    .getRange(GROUP_START_ROW, 1, GROUP_ROWS_COUNT, 1)
    .shiftRowGroupDepth(1);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Группировка квартальных KPI и итогов создана',
    'ACC',
    3
  );
}

/**
 * Собирает все уникальные группировки строк активного листа.
 */
function getAllRowGroupsV1_(sheet) {
  const groups = [];
  const seen = new Set();
  const maxDepth = 8;
  const lastRow = sheet.getLastRow();

  for (let row = 1; row <= lastRow; row++) {
    for (let depth = 1; depth <= maxDepth; depth++) {
      try {
        const group = sheet.getRowGroup(row, depth);
        if (!group) continue;

        const range = group.getRange();
        const startRow = range.getRow();
        const numRows = range.getNumRows();
        const key = `${startRow}:${numRows}:${depth}`;

        if (seen.has(key)) continue;

        seen.add(key);

        groups.push({
          group,
          depth,
          startRow,
          numRows,
          endRow: startRow + numRows - 1,
          header: startRow > 1
            ? String(sheet.getRange(startRow - 1, 1).getDisplayValue()).trim()
            : ''
        });

      } catch (error) {
        // На этой строке и глубине группы нет.
      }
    }
  }

  return groups;
}


/**
 * Классифицирует группировки листа.
 */
function getAccountRowGroupsV1_(sheet) {
  const allGroups = getAllRowGroupsV1_(sheet);

  const sprintGroups = allGroups.filter(item =>
    item.header.startsWith('СПРИНТ №')
  );

  const quarterGroups = allGroups.filter(item =>
    item.header === 'КВАРТАЛЬНЫЕ KPI'
  );

  const monthCandidates = allGroups.filter(item =>
    item.header.startsWith('МЕСЯЦ ')
  );

  const monthGroups = [];
  const monthKpiGroups = [];

  const groupsByStartRow = {};

  monthCandidates.forEach(item => {
    if (!groupsByStartRow[item.startRow]) {
      groupsByStartRow[item.startRow] = [];
    }

    groupsByStartRow[item.startRow].push(item);
  });

  Object.values(groupsByStartRow).forEach(items => {
    items.sort((a, b) => a.numRows - b.numRows);

    // Короткая группа — KPI и итоги месяца.
    monthKpiGroups.push(items[0]);

    // Длинная группа — весь месяц.
    if (items.length > 1) {
      monthGroups.push(items[items.length - 1]);
    }
  });

  return {
    allGroups,
    quarterGroups,
    monthGroups,
    monthKpiGroups,
    sprintGroups
  };
}

//медленные
/**
 * Гарантированно сворачивает все группы всех уровней.
 */
function collapseAllAccountGroupsV1() {
  const sheet = getActiveAccountSheetV1_();
  const { allGroups } = getAccountRowGroupsV1_(sheet);

  allGroups
    .sort((a, b) => b.depth - a.depth)
    .forEach(item => item.group.collapse());
}


/**
 * Гарантированно разворачивает все группы всех уровней.
 */
function expandAllAccountGroupsV1() {
  const sheet = getActiveAccountSheetV1_();
  const { allGroups } = getAccountRowGroupsV1_(sheet);

  allGroups
    .sort((a, b) => a.depth - b.depth)
    .forEach(item => item.group.expand());
}


/**
 * Разворачивает только общие группы месяцев.
 *
 * Внутреннее состояние сохраняется:
 * KPI и спринты остаются свёрнутыми или развёрнутыми
 * так, как были до вызова функции.
 */
function expandMonthsKeepInnerStateV1() {
  const sheet = getActiveAccountSheetV1_();
  const { monthGroups } = getAccountRowGroupsV1_(sheet);

  monthGroups.forEach(item => item.group.expand());
}


/**
 * Показывает только KPI и итоги месяцев.
 *
 * Общие группы месяцев разворачиваются.
 * KPI и итоги месяца разворачиваются.
 * Все спринты остаются свёрнутыми.
 */
function showOnlyMonthKpiV1() {
  const sheet = getActiveAccountSheetV1_();

  const {
    allGroups,
    monthGroups,
    monthKpiGroups
  } = getAccountRowGroupsV1_(sheet);

  // Сначала гарантированно сворачиваем всё.
  allGroups
    .sort((a, b) => b.depth - a.depth)
    .forEach(item => item.group.collapse());

  // Затем открываем только нужные уровни.
  monthGroups.forEach(item => item.group.expand());
  monthKpiGroups.forEach(item => item.group.expand());
}


/**
 * Разворачивает KPI и итоги месяцев,
 * не меняя состояние групп спринтов.
 */
function expandMonthKpiKeepSprintsV1() {
  const sheet = getActiveAccountSheetV1_();

  const {
    monthGroups,
    monthKpiGroups
  } = getAccountRowGroupsV1_(sheet);

  monthGroups.forEach(item => item.group.expand());
  monthKpiGroups.forEach(item => item.group.expand());
}

//быстрые
/**
 * Свернуть абсолютно все группы.
 */
function collapseAllAccountGroupsV1() {
  const sheet = getActiveAccountSheetV1_();

  // Уровень 0 — все группы свёрнуты
  sheet.expandRowGroupsUpToDepth(0);
}


/**
 * Развернуть абсолютно все группы.
 */
function expandAllAccountGroupsV1() {
  const sheet = getActiveAccountSheetV1_();

  // Заведомо больше фактического количества уровней
  sheet.expandRowGroupsUpToDepth(8);
}


/**
 * Развернуть только общие группы месяцев.
 * KPI и спринты внутри остаются свёрнутыми.
 */
function showMonthsOnlyV1() {
  const sheet = getActiveAccountSheetV1_();

  // Открываем только первый уровень
  sheet.expandRowGroupsUpToDepth(1);
}


/**
 * Показать месяцы, KPI и итоги месяца.
 * Спринты остаются свёрнутыми.
 */
function showMonthKpiOnlyV1() {
  const sheet = getActiveAccountSheetV1_();

  // Сначала открываем только месяцы
  sheet.expandRowGroupsUpToDepth(1);

  // Затем точечно открываем KPI и итоги каждого месяца
  const monthKpiRanges = [
    '45:69',
    '153:177',
    '261:285'
  ];

  monthKpiRanges.forEach(a1 => {
    try {
      sheet.getRange(a1).expandGroups();
    } catch (error) {}
  });
}


/**
 * Открыть квартальные KPI и итоги квартала.
 * Остальные группы остаются в текущем состоянии.
 */
function expandQuarterKpiV1() {
  const sheet = getActiveAccountSheetV1_();

  try {
    sheet.getRange('20:41').expandGroups();
  } catch (error) {}
}

/**
 * Рабочий режим.
 *
 * Видны:
 * - квартальные KPI;
 * - месяцы;
 * - KPI месяцев.
 *
 * Спринты скрыты.
 */
function showWorkViewV1() {

  const sheet = getActiveAccountSheetV1_();

  // Открыть месяц
  sheet.expandRowGroupsUpToDepth(1);

  // Открыть квартальные KPI
  try {
    sheet.getRange("20:41").expandGroups();
  } catch (e) {}

  // Открыть KPI месяцев
  [
    "45:69",
    "153:177",
    "261:285"
  ].forEach(r => {
    try {
      sheet.getRange(r).expandGroups();
    } catch (e) {}
  });

}