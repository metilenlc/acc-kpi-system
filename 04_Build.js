/*
Файл шаблона листа
*/

//
const ACC_TEMPLATE_MAIN_FONT_SIZE = LAYOUT.STYLE.FONT.DEFAULT_SIZE;

function buildAccTemplateV1_Layout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ACC_TEMPLATE');

  if (!sheet) {
    throw new Error('Лист ACC_TEMPLATE не найден');
  }

  clearRowGroupsV1_(sheet);

  const DARK = '#0B5D1E';
  const LIGHT = '#EEF7EA';
  const BORDER = '#D9E0D6';
  const WHITE = '#FFFFFF';

  const TOTAL_COLS = 90;
  const TOTAL_ROWS = 520;

  const TITLE_ROW = 1;
  const TITLE_HEIGHT = 36;

  const QUARTER_KPI_START_ROW = 19;
  const MONTH_COUNT = LAYOUT.MONTH.COUNT;

  // Полная очистка листа
  sheet.clear();
  sheet.clearFormats();

  try {
    sheet
      .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .breakApart();
  } catch (e) {}

  sheet
    .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .clearDataValidations();

  // Подготовка размера листа
  if (sheet.getMaxColumns() < TOTAL_COLS) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      TOTAL_COLS - sheet.getMaxColumns()
    );
  }

  if (sheet.getMaxRows() < TOTAL_ROWS) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      TOTAL_ROWS - sheet.getMaxRows()
    );
  }

  sheet.setHiddenGridlines(true);

  for (let c = 1; c <= TOTAL_COLS; c++) {
    sheet.setColumnWidth(c, 20);
  }

  sheet
    .getRange(1, 1, TOTAL_ROWS, TOTAL_COLS)
    .setFontFamily(LAYOUT.STYLE.FONT.FAMILY)
    .setFontSize(ACC_TEMPLATE_MAIN_FONT_SIZE)
    .setBackground(WHITE)
    .setVerticalAlignment('middle')
    .setWrap(true);

  const kpis = [
    'Подписчики',
    'Охват',
    'Переходы по ссылке',
    'Клики Stories',
    'Запросы в Direct',
    'MQL',
    'Стоимость MQL',
    'Стоимость консультации',
    'Выручка новых продаж',
    'Новые клиенты',
    'Пополнение базы'
  ];

  // Заголовок
  mergeText(
    sheet,
    TITLE_ROW,
    1,
    TITLE_ROW,
    TOTAL_COLS,
    'ACC_TEMPLATE V1',
    DARK,
    'white',
    true,
    LAYOUT.STYLE.FONT.SHEET_TITLE_SIZE
  );

  sheet.setRowHeight(TITLE_ROW, TITLE_HEIGHT);

  // Паспорт аккаунта
  buildPassport_(sheet, DARK, LIGHT, BORDER);

  // KPI квартала
  buildKpiBlock_(
    sheet,
    QUARTER_KPI_START_ROW,
    'КВАРТАЛЬНЫЕ KPI',
    kpis,
    DARK,
    LIGHT,
    BORDER,
    'План квартал',
    'Факт квартал',
    LAYOUT.STYLE.FONT.QUARTER_MONTH_TITLE_SIZE
  );

  // Итоги квартала под квартальными KPI
  const quarterSummaryEndRow = buildQuarterSummary_(
    sheet,
    QUARTER_KPI_START_ROW + 15,
    DARK,
    LIGHT,
    BORDER
  );

  // Месяцы
  for (let monthNumber = 1; monthNumber <= MONTH_COUNT; monthNumber++) {
    const monthLayout = getMonthLayoutV1_(monthNumber - 1);

    buildMonthBlock_(
      sheet,
      monthLayout.monthRow,
      monthNumber,
      kpis,
      DARK,
      LIGHT,
      BORDER
    );
  }

  applyTemplateSpacerRowHeightsV1_(sheet);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(0);

  setupQuarterKpiSummaryGroupV1(sheet, true);
  setupMonthGroupsV1(sheet, true);
  setupMonthKpiSummaryGroupsV1(sheet, true);
  setupSprintRowGroupsV1(sheet, true);
  sheet.expandRowGroupsUpToDepth(1);

  SpreadsheetApp.flush();

  ss.toast(
    'ACC_TEMPLATE пересобран',
    'ACC',
    3
  );
}

//
function buildPassport_(sheet, DARK, LIGHT, BORDER) {
  mergeText(sheet, 3, 1, 3, 40, 'ПАСПОРТ АККАУНТА', DARK, 'white', true, LAYOUT.STYLE.FONT.SECTION_TITLE_SIZE);

  const fields = ['Аккаунт', 'Соцсеть', 'Ответственный', 'Год', 'Квартал', 'Статус'];

  for (let i = 0; i < fields.length; i++) {
    const r = 4 + i;

    mergeText(sheet, r, 1, r, 12, fields[i], LIGHT, 'black', true, ACC_TEMPLATE_MAIN_FONT_SIZE);
    mergeEmpty(sheet, r, 13, r, 40);
    if (i < 4) {
      sheet
        .getRange(r, 13, 1, 28)
        .setHorizontalAlignment('left');
    }
    sheet.setRowHeight(r, 28);
  }

  sheet.getRange(3, 1, 7, 40)
    .setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID);

  mergeText(sheet, 3, 42, 3, 90, 'ЦЕЛЬ КВАРТАЛА', DARK, 'white', true, LAYOUT.STYLE.FONT.SECTION_TITLE_SIZE);
  mergeEmpty(sheet, 4, 42, 10, 90);
  sheet.getRange(4, 42, 7, 49).setVerticalAlignment('top');

  sheet.getRange(3, 42, 8, 49)
    .setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID);

  mergeText(sheet, 12, 42, 12, 90, 'КОММЕНТАРИИ', DARK, 'white', true, ACC_TEMPLATE_MAIN_FONT_SIZE);
  mergeEmpty(sheet, 13, 42, 15, 90);
  sheet.getRange(13, 42, 3, 49).setVerticalAlignment('top');

  sheet.getRange(12, 42, 4, 49)
    .setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(17, 1, 1, 90).merge().setBackground(DARK);
  sheet.setRowHeight(17, LAYOUT.STYLE.ROW_HEIGHTS.THIN);
}

//
function buildKpiBlock_(sheet, startRow, title, kpis, DARK, LIGHT, BORDER, planTitle, factTitle, titleFontSize) {

  const TOTAL_COLS = 90;
  const KPI_ROWS = LAYOUT.MONTH.KPI_COUNT;

  const COL_KPI_START = 1;
  const COL_KPI_END = 32;

  const COL_PLAN_START = 33;
  const COL_PLAN_END = 46;

  const COL_FACT_START = 47;
  const COL_FACT_END = 60;

  const COL_PERCENT_START = 61;
  const COL_PERCENT_END = 64;

  const COL_BAR_START = 65;
  const COL_BAR_END = 84;
  const COL_BAR_WIDTH = COL_BAR_END - COL_BAR_START + 1;

  const COL_RESERVE_START = 85;
  const COL_RESERVE_END = 90;

  const HEADER_HEIGHT = 32;
  const ROW_HEIGHT = 40;

  mergeText(sheet, startRow, 1, startRow, TOTAL_COLS, title, DARK, 'white', true, titleFontSize || ACC_TEMPLATE_MAIN_FONT_SIZE);
  sheet.setRowHeight(startRow, HEADER_HEIGHT);

  mergeText(sheet, startRow + 1, COL_KPI_START, startRow + 1, COL_KPI_END, 'KPI', LIGHT, 'black', true, ACC_TEMPLATE_MAIN_FONT_SIZE);
  mergeText(sheet, startRow + 1, COL_PLAN_START, startRow + 1, COL_PLAN_END, planTitle, LIGHT, 'black', true, ACC_TEMPLATE_MAIN_FONT_SIZE);
  mergeText(sheet, startRow + 1, COL_FACT_START, startRow + 1, COL_FACT_END, factTitle, LIGHT, 'black', true, ACC_TEMPLATE_MAIN_FONT_SIZE);
  mergeText(sheet, startRow + 1, COL_PERCENT_START, startRow + 1, COL_PERCENT_END, '%', LIGHT, 'black', true, ACC_TEMPLATE_MAIN_FONT_SIZE);
  mergeText(sheet, startRow + 1, COL_BAR_START, startRow + 1, COL_BAR_END, 'Прогресс', LIGHT, 'black', true, ACC_TEMPLATE_MAIN_FONT_SIZE);
  mergeEmpty(sheet, startRow + 1, COL_RESERVE_START, startRow + 1, COL_RESERVE_END);

  sheet.setRowHeight(startRow + 1, HEADER_HEIGHT);

  for (let i = 0; i < KPI_ROWS; i++) {
    const r = startRow + 2 + i;

    mergeEmpty(sheet, r, COL_KPI_START, r, COL_KPI_END);
    mergeEmpty(sheet, r, COL_PLAN_START, r, COL_PLAN_END);
    mergeEmpty(sheet, r, COL_FACT_START, r, COL_FACT_END);
    mergeEmpty(sheet, r, COL_PERCENT_START, r, COL_PERCENT_END);
    mergeEmpty(sheet, r, COL_BAR_START, r, COL_BAR_END);
    mergeEmpty(sheet, r, COL_RESERVE_START, r, COL_RESERVE_END);

    const percentCell = sheet.getRange(r, COL_PERCENT_START);
    const barRange = sheet.getRange(r, COL_BAR_START, 1, COL_BAR_WIDTH);

    insertProgressBarV1_(sheet, percentCell, barRange);

    sheet.setRowHeight(r, ROW_HEIGHT);
  }

  sheet.getRange(startRow + 1, 1, KPI_ROWS + 1, TOTAL_COLS)
    .setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

//
function buildMonthBlock_(
    sheet,
    startRow,
    monthNumber,
    kpis,
    DARK,
    LIGHT,
    BORDER
  ) {
  const TOTAL_COLS = 90;
  const MONTH_TITLE_HEIGHT = 35;

  const monthIndex = monthNumber - 1;
  const monthLayout = getMonthLayoutV1_(monthIndex);

  mergeText(
    sheet,
    startRow,
    1,
    startRow,
    TOTAL_COLS,
    'МЕСЯЦ ' + monthNumber,
    DARK,
    'white',
    true,
    LAYOUT.STYLE.FONT.QUARTER_MONTH_TITLE_SIZE
  );

  sheet.setRowHeight(startRow, MONTH_TITLE_HEIGHT);

  buildKpiBlock_(
    sheet,
    monthLayout.kpiRow,
    'KPI МЕСЯЦА',
    kpis,
    DARK,
    LIGHT,
    BORDER,
    'План',
    'Факт'
  );

  buildMonthSummary_(
    sheet,
    monthLayout.summaryRow,
    DARK,
    LIGHT,
    BORDER
  );

  for (
    let sprintIndex = 0;
    sprintIndex < LAYOUT.SPRINT.COUNT;
    sprintIndex++
  ) {
    const sprintLayout = getSprintLayoutV1_(
      monthIndex,
      sprintIndex
    );

    buildSprintCard_(
      sheet,
      sprintLayout.sprintTop,
      monthNumber,
      sprintIndex + 1,
      DARK,
      LIGHT,
      BORDER
    );
  }

  const lastSprint = getSprintLayoutV1_(
    monthIndex,
    LAYOUT.SPRINT.COUNT - 1
  );

  return lastSprint.bottomGapRow;
}

/**
 * ==========================================
 * buildSprintCard_
 * ==========================================
 *
 * Назначение:
 * Строит один блок спринта в ACC_TEMPLATE.
 *
 * Создаёт:
 * • Заголовок спринта
 * • Выбор спринта
 * • Даты начала и окончания
 * • ProgressBar спринта
 * • Таблицу гипотез
 * • Итоги спринта
 */
function buildSprintCard_(
  sheet,
  startRow,
  monthNumber,
  sprintNumber,
  DARK,
  LIGHT,
  BORDER
) {
  const TOTAL_COLS = 90;

  const rowHeights = LAYOUT.SPRINT.ROW_HEIGHTS;
  const rowOffsets = LAYOUT.SPRINT.ROW_OFFSETS;

  const COL_SPRINT_LABEL_START = 1;
  const COL_SPRINT_LABEL_END = 10;
  const COL_SPRINT_VALUE_START = 11;
  const COL_SPRINT_VALUE_END = 36;

  const COL_PROGRESS_LABEL_START = 1;
  const COL_PROGRESS_LABEL_END = 10;
  const COL_PROGRESS_PERCENT_START = 11;
  const COL_PROGRESS_PERCENT_END = 14;
  const COL_PROGRESS_BAR_START = 15;
  const COL_PROGRESS_BAR_END = 36;

  const COL_DATE_LABEL_START = 49;
  const COL_DATE_LABEL_END = 57;
  const COL_DATE_VALUE_START = 58;
  const COL_DATE_VALUE_END = 72;

  const HYP_WIDTHS = [
    2,
    24,
    16,
    5,
    5,
    5,
    5,
    12,
    8,
    8
  ];

  const HYP_HEADERS = [
    '✓',
    'Гипотеза',
    'KPI',
    'План',
    'Факт',
    'Ожид.',
    'Факт.эфф.',
    'Исполнитель',
    'Статус',
    'Результат'
  ];

  const summaryConfig = [
    {
      title: 'Заполнено',
      range: LAYOUT.SPRINT_SUMMARY.FILLED
    },
    {
      title: 'Выполнено',
      range: LAYOUT.SPRINT_SUMMARY.DONE
    },
    {
      title: 'В работе',
      range: LAYOUT.SPRINT_SUMMARY.IN_PROGRESS
    },
    {
      title: 'Отложено',
      range: LAYOUT.SPRINT_SUMMARY.POSTPONED
    },
    {
      title: 'Прогресс',
      range: LAYOUT.SPRINT_SUMMARY.PROGRESS
    },
    {
      title: 'Успешно',
      range: LAYOUT.SPRINT_SUMMARY.SUCCESS
    },
    {
      title: 'Частично успешно',
      range: LAYOUT.SPRINT_SUMMARY.PARTIAL
    },
    {
      title: 'Неуспешно',
      range: LAYOUT.SPRINT_SUMMARY.FAILED
    }
  ];

  // Заголовок спринта
  mergeText(
    sheet,
    startRow,
    1,
    startRow,
    TOTAL_COLS,
    `СПРИНТ ${monthNumber}.${sprintNumber}`,
    DARK,
    'white',
    true,
    LAYOUT.STYLE.FONT.SPRINT_TITLE_SIZE
  );

  sheet.setRowHeight(
    startRow,
    rowHeights.TITLE
  );

  // Выбор спринта и дата начала
  const selectorRow = startRow + rowOffsets.SELECTOR;

  mergeText(
    sheet,
    selectorRow,
    COL_SPRINT_LABEL_START,
    selectorRow,
    COL_SPRINT_LABEL_END,
    'Спринт',
    LIGHT,
    'black',
    true,
    ACC_TEMPLATE_MAIN_FONT_SIZE
  );

  mergeEmpty(
    sheet,
    selectorRow,
    COL_SPRINT_VALUE_START,
    selectorRow,
    COL_SPRINT_VALUE_END
  );

  setInputBorder_(
    sheet,
    selectorRow,
    COL_SPRINT_VALUE_START,
    COL_SPRINT_VALUE_END,
    BORDER
  );

  mergeText(
    sheet,
    selectorRow,
    COL_DATE_LABEL_START,
    selectorRow,
    COL_DATE_LABEL_END,
    'Дата начала',
    LIGHT,
    'black',
    true,
    ACC_TEMPLATE_MAIN_FONT_SIZE
  );

  mergeEmpty(
    sheet,
    selectorRow,
    COL_DATE_VALUE_START,
    selectorRow,
    COL_DATE_VALUE_END
  );

  setInputBorder_(
    sheet,
    selectorRow,
    COL_DATE_VALUE_START,
    COL_DATE_VALUE_END,
    BORDER
  );

  sheet.setRowHeight(
    selectorRow,
    rowHeights.SELECTOR
  );

  // Разделитель
  sheet.setRowHeight(
    startRow + rowOffsets.SEPARATOR,
    LAYOUT.STYLE.ROW_HEIGHTS.THIN
  );

  // Прогресс и дата окончания
  const progressRow = startRow + rowOffsets.PROGRESS;

  mergeText(
    sheet,
    progressRow,
    COL_PROGRESS_LABEL_START,
    progressRow,
    COL_PROGRESS_LABEL_END,
    'Прогресс',
    LIGHT,
    'black',
    true,
    ACC_TEMPLATE_MAIN_FONT_SIZE
  );

  mergeEmpty(
    sheet,
    progressRow,
    COL_PROGRESS_PERCENT_START,
    progressRow,
    COL_PROGRESS_PERCENT_END
  );

  setInputBorder_(
    sheet,
    progressRow,
    COL_PROGRESS_PERCENT_START,
    COL_PROGRESS_PERCENT_END,
    BORDER
  );

  mergeEmpty(
    sheet,
    progressRow,
    COL_PROGRESS_BAR_START,
    progressRow,
    COL_PROGRESS_BAR_END
  );

  setInputBorder_(
    sheet,
    progressRow,
    COL_PROGRESS_BAR_START,
    COL_PROGRESS_BAR_END,
    '#000000'
  );

  mergeText(
    sheet,
    progressRow,
    COL_DATE_LABEL_START,
    progressRow,
    COL_DATE_LABEL_END,
    'Дата окончания',
    LIGHT,
    'black',
    true,
    ACC_TEMPLATE_MAIN_FONT_SIZE
  );

  mergeEmpty(
    sheet,
    progressRow,
    COL_DATE_VALUE_START,
    progressRow,
    COL_DATE_VALUE_END
  );

  setInputBorder_(
    sheet,
    progressRow,
    COL_DATE_VALUE_START,
    COL_DATE_VALUE_END,
    BORDER
  );

  sheet.setRowHeight(
    progressRow,
    rowHeights.PROGRESS
  );

  // Заголовок таблицы гипотез
  const hypHeaderRow =
    startRow + rowOffsets.HYPOTHESIS_HEADER;
  let currentCol = 1;

  HYP_HEADERS.forEach((header, index) => {
    const width = HYP_WIDTHS[index];

    mergeText(
      sheet,
      hypHeaderRow,
      currentCol,
      hypHeaderRow,
      currentCol + width - 1,
      header,
      LIGHT,
      'black',
      true,
      ACC_TEMPLATE_MAIN_FONT_SIZE
    );

    currentCol += width;
  });

  sheet.setRowHeight(
    hypHeaderRow,
    rowHeights.HYPOTHESIS_HEADER
  );

  // Строки гипотез
  const firstHypRow = hypHeaderRow + 1;
  const lastHypRow =
    firstHypRow +
    LAYOUT.SPRINT.HYPOTHESIS_COUNT -
    1;

  for (
    let row = firstHypRow;
    row <= lastHypRow;
    row++
  ) {
    currentCol = 1;

    HYP_WIDTHS.forEach(width => {
      mergeEmpty(
        sheet,
        row,
        currentCol,
        row,
        currentCol + width - 1
      );

      currentCol += width;
    });

    sheet.setRowHeight(
      row,
      rowHeights.HYPOTHESIS
    );
  }

  sheet
    .getRange(
      hypHeaderRow,
      1,
      LAYOUT.SPRINT.HYPOTHESIS_COUNT + 1,
      TOTAL_COLS
    )
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      BORDER,
      SpreadsheetApp.BorderStyle.SOLID
    );

  // Отступ перед итогами
  const summaryOffsets = LAYOUT.SPRINT.SUMMARY_OFFSETS;
  const summaryGapRow =
    lastHypRow + summaryOffsets.GAP;

  sheet.setRowHeight(
    summaryGapRow,
    LAYOUT.STYLE.ROW_HEIGHTS.THIN
  );

  // Итоги спринта
  const summaryHeaderRow =
    lastHypRow + summaryOffsets.HEADER;
  const summaryValueRow =
    lastHypRow + summaryOffsets.VALUES;

  summaryConfig.forEach(item => {
    mergeText(
      sheet,
      summaryHeaderRow,
      item.range.START_COLUMN,
      summaryHeaderRow,
      item.range.END_COLUMN,
      item.title,
      LIGHT,
      'black',
      true,
      ACC_TEMPLATE_MAIN_FONT_SIZE
    );

    mergeEmpty(
      sheet,
      summaryValueRow,
      item.range.START_COLUMN,
      summaryValueRow,
      item.range.END_COLUMN
    );

    setInputBorder_(
      sheet,
      summaryValueRow,
      item.range.START_COLUMN,
      item.range.END_COLUMN,
      BORDER
    );

    sheet
      .getRange(
        summaryValueRow,
        item.range.START_COLUMN,
        1,
        item.range.END_COLUMN -
          item.range.START_COLUMN +
          1
      )
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });

  sheet.setRowHeight(
    summaryHeaderRow,
    rowHeights.SUMMARY_HEADER
  );

  sheet.setRowHeight(
    summaryValueRow,
    rowHeights.SUMMARY_VALUE
  );

  // Пустая строка между спринтами
  sheet.setRowHeight(
    lastHypRow + summaryOffsets.BOTTOM,
    LAYOUT.STYLE.ROW_HEIGHTS.EMPTY
  );
}

function buildMonthSummary_(sheet, startRow, DARK, LIGHT, BORDER) {
  return buildSummaryBlock_(
    sheet,
    startRow,
    'ИТОГИ МЕСЯЦА',
    'Прогресс месяца',
    DARK,
    LIGHT,
    BORDER
  );
}

function buildQuarterSummary_(sheet, startRow, DARK, LIGHT, BORDER) {
  return buildSummaryBlock_(
    sheet,
    startRow,
    'ИТОГИ КВАРТАЛА',
    'Прогресс квартала',
    DARK,
    LIGHT,
    BORDER
  );
}

function buildSummaryBlock_(sheet, startRow, title, progressTitle, DARK, LIGHT, BORDER) {

  const TOTAL_COLS = 90;

  const TITLE_HEIGHT = 26;
  const DATA_ROW_HEIGHT = 28;
  const GAP_HEIGHT = LAYOUT.STYLE.ROW_HEIGHTS.EMPTY;

  const COL_LEFT_LABEL_START = 8;
  const COL_LEFT_LABEL_END = 18;
  const COL_LEFT_VALUE_START = 19;
  const COL_LEFT_VALUE_END = 22;

  const COL_RIGHT_LABEL_START = 27;
  const COL_RIGHT_LABEL_END = 38;
  const COL_RIGHT_VALUE_START = 39;
  const COL_RIGHT_VALUE_END = 42;

  const COL_PROGRESS_LABEL_START = 8;
  const COL_PROGRESS_LABEL_END = 18;
  const COL_PROGRESS_PERCENT_START = 19;
  const COL_PROGRESS_PERCENT_END = 22;
  const COL_PROGRESS_BAR_START = 23;
  const COL_PROGRESS_BAR_END = 42;

  mergeText(sheet, startRow, 1, startRow, TOTAL_COLS, title, DARK, 'white', true, LAYOUT.STYLE.FONT.SECTION_TITLE_SIZE);
  sheet.setRowHeight(startRow, TITLE_HEIGHT);

  sheet.setRowHeight(startRow + 1, GAP_HEIGHT);

  const leftItems = [
    'Всего гипотез',
    'Выполнено',
    'Конверсия'
  ];

  const rightItems = [
    'Успешно',
    'Частично успешно',
    'Неуспешно'
  ];

  for (let i = 0; i < 3; i++) {
    const r = startRow + 2 + i;

    mergeText(sheet, r, COL_LEFT_LABEL_START, r, COL_LEFT_LABEL_END, leftItems[i], LIGHT, 'black', true, ACC_TEMPLATE_MAIN_FONT_SIZE);
    mergeEmpty(sheet, r, COL_LEFT_VALUE_START, r, COL_LEFT_VALUE_END);
    setInputBorder_(sheet, r, COL_LEFT_VALUE_START, COL_LEFT_VALUE_END, BORDER);

    mergeText(sheet, r, COL_RIGHT_LABEL_START, r, COL_RIGHT_LABEL_END, rightItems[i], LIGHT, 'black', true, ACC_TEMPLATE_MAIN_FONT_SIZE);
    mergeEmpty(sheet, r, COL_RIGHT_VALUE_START, r, COL_RIGHT_VALUE_END);
    setInputBorder_(sheet, r, COL_RIGHT_VALUE_START, COL_RIGHT_VALUE_END, BORDER);

    sheet.setRowHeight(r, DATA_ROW_HEIGHT);
  }

  sheet.setRowHeight(startRow + 5, GAP_HEIGHT);

  mergeText(sheet, startRow + 6, COL_PROGRESS_LABEL_START, startRow + 6, COL_PROGRESS_LABEL_END, progressTitle, LIGHT, 'black', true, ACC_TEMPLATE_MAIN_FONT_SIZE);
  mergeEmpty(sheet, startRow + 6, COL_PROGRESS_PERCENT_START, startRow + 6, COL_PROGRESS_PERCENT_END);
  setInputBorder_(sheet, startRow + 6, COL_PROGRESS_PERCENT_START, COL_PROGRESS_PERCENT_END, BORDER);

  mergeEmpty(sheet, startRow + 6, COL_PROGRESS_BAR_START, startRow + 6, COL_PROGRESS_BAR_END);
  setInputBorder_(sheet, startRow + 6, COL_PROGRESS_BAR_START, COL_PROGRESS_BAR_END, '#000000');

  sheet.setRowHeight(startRow + 6, DATA_ROW_HEIGHT);

  return startRow + 8;
}

function applyTemplateSpacerRowHeightsV1_(sheet) {
  const emptyHeight = LAYOUT.STYLE.ROW_HEIGHTS.EMPTY;
  const thinHeight = LAYOUT.STYLE.ROW_HEIGHTS.THIN;
  const quarterKpiTitleRow = LAYOUT.QUARTER.KPI.ROW - 1;
  const emptyRows = new Set([
    LAYOUT.SHEET.FIRST_ROW + 1,
    LAYOUT.PASSPORT.LAST_ROW + 2,
    quarterKpiTitleRow - 3,
    quarterKpiTitleRow - 1,
  ]);
  const thinRows = new Set([quarterKpiTitleRow - 2]);

  const quarterKpiBottomGapRow =
    LAYOUT.QUARTER.KPI.ROW + LAYOUT.QUARTER.KPI_COUNT + 1;
  const quarterSummaryTitleRow = LAYOUT.QUARTER.SUMMARY.ROW - 2;

  emptyRows.add(quarterKpiBottomGapRow);
  emptyRows.add(quarterSummaryTitleRow + 1);
  emptyRows.add(quarterSummaryTitleRow + 5);

  for (
    let row = quarterSummaryTitleRow + 7;
    row < LAYOUT.MONTH.FIRST_ROW;
    row++
  ) {
    emptyRows.add(row);
  }

  for (let monthIndex = 0; monthIndex < LAYOUT.MONTH.COUNT; monthIndex++) {
    const month = getMonthLayoutV1_(monthIndex);
    const lastSprint = getSprintLayoutV1_(
      monthIndex,
      LAYOUT.SPRINT.COUNT - 1
    );

    emptyRows.add(month.titleRow + 1);
    emptyRows.add(month.kpiBottomGapRow);
    emptyRows.add(month.summaryRow + 1);
    emptyRows.add(month.summaryRow + 5);

    for (
      let row = month.summaryRow + 7;
      row < month.firstSprintRow;
      row++
    ) {
      emptyRows.add(row);
    }

    for (let sprintIndex = 0; sprintIndex < LAYOUT.SPRINT.COUNT; sprintIndex++) {
      const sprint = getSprintLayoutV1_(monthIndex, sprintIndex);

      emptyRows.add(sprint.sprintTop + 1);
      emptyRows.add(sprint.headerRow - 1);
      emptyRows.add(sprint.bottomGapRow);

      thinRows.add(
        sprint.sprintTop + LAYOUT.SPRINT.ROW_OFFSETS.SEPARATOR
      );
      thinRows.add(sprint.summaryGapRow);
    }

    const monthEndRow = month.titleRow + getMonthHeightV1_() - 1;
    for (let row = lastSprint.bottomGapRow + 1; row <= monthEndRow; row++) {
      emptyRows.add(row);
    }
  }

  emptyRows.forEach(row => sheet.setRowHeight(row, emptyHeight));
  thinRows.forEach(row => sheet.setRowHeight(row, thinHeight));
}
