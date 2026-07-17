/*
==================================================
10_Groups.gs
Группировки строк
==================================================
*/

function getRowGroupTargetSheetV1_(targetSheet) {
  if (targetSheet) return targetSheet;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeName = activeSheet && activeSheet.getName();

  if (activeName && activeName.startsWith('ACC_')) {
    return activeSheet;
  }

  const template = ss.getSheetByName('ACC_TEMPLATE');
  if (!template) throw new Error('Лист ACC_TEMPLATE не найден.');

  return template;
}

/**
 * ==========================================
 * setupSprintRowGroupsV1
 * ==========================================
 *
 * Назначение:
 * Создаёт внутренние группировки для спринтов.
 *
 * Логика:
 * • Заголовок СПРИНТ M.N остаётся видимым.
 * • Тело спринта группируется.
 * • Каждый спринт можно свернуть отдельно.
 */
function setupSprintRowGroupsV1(targetSheet, silent) {
  const sheet = getRowGroupTargetSheetV1_(targetSheet);

  for (let monthIndex = 0; monthIndex < LAYOUT.MONTH.COUNT; monthIndex++) {
    for (let sprintIndex = 0; sprintIndex < LAYOUT.SPRINT.COUNT; sprintIndex++) {
      const sprint = getSprintLayoutV1_(monthIndex, sprintIndex);
      const groupStartRow = sprint.sprintTop + 1;
      const groupRowsCount = sprint.totalsRow - sprint.sprintTop;

      sheet
        .getRange(groupStartRow, 1, groupRowsCount, 1)
        .shiftRowGroupDepth(1);
    }
  }

  if (!silent) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Группировки спринтов созданы',
      'ACC',
      3
    );
  }
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
function clearRowGroupsV1_(targetSheet) {
  const sheet = getRowGroupTargetSheetV1_(targetSheet);
  const maxDepth = 8;
  const allRows = sheet.getRange(1, 1, sheet.getMaxRows(), 1);

  try {
    sheet.expandAllRowGroups();
  } catch (error) {}

  for (let pass = 0; pass < maxDepth; pass++) {
    allRows.shiftRowGroupDepth(-1);
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
function setupMonthKpiSummaryGroupsV1(targetSheet, silent) {
  const sheet = getRowGroupTargetSheetV1_(targetSheet);

  for (let monthIndex = 0; monthIndex < LAYOUT.MONTH.COUNT; monthIndex++) {
    const month = getMonthLayoutV1_(monthIndex);

    shiftRowGroupRangeV1_(
      sheet,
      month.titleRow + 1,
      month.firstSprintRow - 3
    );
  }

  if (!silent) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Группировки KPI и итогов месяца созданы',
      'ACC',
      3
    );
  }
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
function setupMonthGroupsV1(targetSheet, silent) {

  const sheet = getRowGroupTargetSheetV1_(targetSheet);

  for (
    let monthIndex = 0;
    monthIndex < LAYOUT.MONTH.COUNT;
    monthIndex++
  ) {
    const month = getMonthLayoutV1_(monthIndex);
    const groupStart = month.monthRow + 1;
    const groupEnd = month.monthRow + getMonthHeightV1_() - 3;
    const groupRows = groupEnd - groupStart + 1;

    sheet
      .getRange(groupStart, 1, groupRows, 1)
      .shiftRowGroupDepth(1);
  }

  if (!silent) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Группировки месяцев созданы',
      'ACC',
      3
    );
  }
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
 * Диапазон рассчитывается от строки под заголовком квартальных KPI
 * до трёх строк перед заголовком первого месяца.
 */
function setupQuarterKpiSummaryGroupV1(targetSheet, silent) {
  const sheet = getRowGroupTargetSheetV1_(targetSheet);
  const quarterTitleRow = LAYOUT.QUARTER.KPI.ROW - 1;
  const groupStartRow = quarterTitleRow + 1;
  const groupEndRow = LAYOUT.MONTH.FIRST_ROW - 4;

  shiftRowGroupRangeV1_(sheet, groupStartRow, groupEndRow);

  if (!silent) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Группировка квартальных KPI и итогов создана',
      'ACC',
      3
    );
  }
}

function shiftRowGroupRangeV1_(sheet, startRow, endRow) {
  if (endRow < startRow) return;

  sheet
    .getRange(startRow, 1, endRow - startRow + 1, 1)
    .shiftRowGroupDepth(1);
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
  const findGroup = (startRow, endRow) => allGroups.find(item =>
    item.startRow === startRow && item.endRow === endRow
  );
  const quarterGroups = [findGroup(
    LAYOUT.QUARTER.KPI.ROW,
    LAYOUT.MONTH.FIRST_ROW - 4
  )].filter(Boolean);
  const monthGroups = [];
  const monthKpiGroups = [];
  const sprintGroups = [];

  for (let monthIndex = 0; monthIndex < LAYOUT.MONTH.COUNT; monthIndex++) {
    const month = getMonthLayoutV1_(monthIndex);
    const monthEnd = month.monthRow + getMonthHeightV1_() - 3;

    const monthGroup = findGroup(month.monthRow + 1, monthEnd);
    if (monthGroup) monthGroups.push(monthGroup);

    const kpiGroup = findGroup(month.titleRow + 1, month.firstSprintRow - 3);
    if (kpiGroup) monthKpiGroups.push(kpiGroup);

    for (let sprintIndex = 0; sprintIndex < LAYOUT.SPRINT.COUNT; sprintIndex++) {
      const sprint = getSprintLayoutV1_(monthIndex, sprintIndex);
      const sprintGroup = findGroup(sprint.sprintTop + 1, sprint.totalsRow);
      if (sprintGroup) sprintGroups.push(sprintGroup);
    }
  }

  return {
    allGroups,
    quarterGroups,
    monthGroups,
    monthKpiGroups,
    sprintGroups
  };
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
  const sheet = getRowGroupTargetSheetV1_();

  // Уровень 0 — все группы свёрнуты
  sheet.expandRowGroupsUpToDepth(0);
}


/**
 * Развернуть абсолютно все группы.
 */
function expandAllAccountGroupsV1() {
  const sheet = getRowGroupTargetSheetV1_();

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
  for (
    let monthIndex = 0;
    monthIndex < LAYOUT.MONTH.COUNT;
    monthIndex++
  ) {
    const month = getMonthLayoutV1_(monthIndex);
    const startRow = month.monthRow + 1;
    const endRow = month.firstSprintRow - 2;

    try {
      sheet
        .getRange(startRow, 1, endRow - startRow + 1, 1)
        .expandGroups();
    } catch (error) {}
  }
}


/**
 * Открыть квартальные KPI и итоги квартала.
 * Остальные группы остаются в текущем состоянии.
 */
function expandQuarterKpiV1() {
  const sheet = getActiveAccountSheetV1_();

  try {
    sheet
      .getRange(
        LAYOUT.QUARTER.KPI.ROW,
        1,
        LAYOUT.MONTH.FIRST_ROW - LAYOUT.QUARTER.KPI.ROW,
        1
      )
      .expandGroups();
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
    sheet
      .getRange(
        LAYOUT.QUARTER.KPI.ROW,
        1,
        LAYOUT.MONTH.FIRST_ROW - LAYOUT.QUARTER.KPI.ROW,
        1
      )
      .expandGroups();
  } catch (e) {}

  // Открыть KPI месяцев
  for (
    let monthIndex = 0;
    monthIndex < LAYOUT.MONTH.COUNT;
    monthIndex++
  ) {
    const month = getMonthLayoutV1_(monthIndex);
    const startRow = month.monthRow + 1;
    const endRow = month.firstSprintRow - 2;

    try {
      sheet
        .getRange(startRow, 1, endRow - startRow + 1, 1)
        .expandGroups();
    } catch (e) {}
  }

}
