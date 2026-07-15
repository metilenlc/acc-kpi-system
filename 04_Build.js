/*
Файл шаблона листа
*/

//
function buildAccTemplateV1_Layout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ACC_TEMPLATE');

  if (!sheet) {
    throw new Error('Лист ACC_TEMPLATE не найден');
  }

  const DARK = '#0B5D1E';
  const LIGHT = '#EEF7EA';
  const BORDER = '#D9E0D6';
  const WHITE = '#FFFFFF';

  const TOTAL_COLS = 90;
  const TOTAL_ROWS = 520;

  const TITLE_ROW = 1;
  const TITLE_HEIGHT = 36;

  const QUARTER_KPI_START_ROW = 19;
  const MONTH_1_START_ROW = 34;
  const MONTH_COUNT = 3;
  const MONTH_GAP_ROWS = 3;

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
    .setFontFamily('Arial')
    .setFontSize(10)
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
    16
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
    'Факт квартал'
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
  let monthStart = quarterSummaryEndRow + 2;

  for (let monthNumber = 1; monthNumber <= MONTH_COUNT; monthNumber++) {
    monthStart = buildMonthBlock_(
      sheet,
      monthStart,
      monthNumber,
      kpis,
      DARK,
      LIGHT,
      BORDER
    );

    monthStart += MONTH_GAP_ROWS;
  }

    sheet.setFrozenRows(3);

  SpreadsheetApp.flush();

  ss.toast(
    'ACC_TEMPLATE пересобран',
    'ACC',
    3
  );
}

//
function buildPassport_(sheet, DARK, LIGHT, BORDER) {
  mergeText(sheet, 3, 1, 3, 40, 'ПАСПОРТ АККАУНТА', DARK, 'white', true, 10);

  const fields = ['Аккаунт', 'Соцсеть', 'Ответственный', 'Год', 'Квартал', 'Статус'];

  for (let i = 0; i < fields.length; i++) {
    const r = 4 + i;

    mergeText(sheet, r, 1, r, 12, fields[i], LIGHT, 'black', true, 10);
    mergeEmpty(sheet, r, 13, r, 40);
    sheet.setRowHeight(r, 28);
  }

  sheet.getRange(3, 1, 7, 40)
    .setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID);

  mergeText(sheet, 3, 42, 3, 90, 'ЦЕЛЬ КВАРТАЛА', DARK, 'white', true, 10);
  mergeEmpty(sheet, 4, 42, 10, 90);
  sheet.getRange(4, 42, 7, 49).setVerticalAlignment('top');

  sheet.getRange(3, 42, 8, 49)
    .setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID);

  mergeText(sheet, 12, 42, 12, 90, 'КОММЕНТАРИИ', DARK, 'white', true, 10);
  mergeEmpty(sheet, 13, 42, 15, 90);
  sheet.getRange(13, 42, 3, 49).setVerticalAlignment('top');

  sheet.getRange(12, 42, 4, 49)
    .setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(17, 1, 1, 90).merge().setBackground(DARK);
  sheet.setRowHeight(17, 10);
}

//
function buildKpiBlock_(sheet, startRow, title, kpis, DARK, LIGHT, BORDER, planTitle, factTitle) {

  const TOTAL_COLS = 90;
  const KPI_ROWS = 12;

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

  mergeText(sheet, startRow, 1, startRow, TOTAL_COLS, title, DARK, 'white', true, 10);
  sheet.setRowHeight(startRow, HEADER_HEIGHT);

  mergeText(sheet, startRow + 1, COL_KPI_START, startRow + 1, COL_KPI_END, 'KPI', LIGHT, 'black', true, 10);
  mergeText(sheet, startRow + 1, COL_PLAN_START, startRow + 1, COL_PLAN_END, planTitle, LIGHT, 'black', true, 10);
  mergeText(sheet, startRow + 1, COL_FACT_START, startRow + 1, COL_FACT_END, factTitle, LIGHT, 'black', true, 10);
  mergeText(sheet, startRow + 1, COL_PERCENT_START, startRow + 1, COL_PERCENT_END, '%', LIGHT, 'black', true, 10);
  mergeText(sheet, startRow + 1, COL_BAR_START, startRow + 1, COL_BAR_END, 'Прогресс', LIGHT, 'black', true, 10);
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
    14
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

  return lastSprint.bottomGapRow + 3;
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
  sprintNumber,
  DARK,
  LIGHT,
  BORDER
) {
  const TOTAL_COLS = 90;

  const TITLE_ROW_HEIGHT = 32;
  const HEADER_ROW_HEIGHT = 28;
  const SEPARATOR_ROW_HEIGHT = 6;
  const HYP_HEADER_ROW_HEIGHT = 32;
  const HYP_ROW_HEIGHT = 48;
  const SUMMARY_TOP_GAP_HEIGHT = 6;
  const SUMMARY_HEADER_ROW_HEIGHT = 26;
  const SUMMARY_VALUE_ROW_HEIGHT = 28;
  const BOTTOM_GAP_HEIGHT = 28;

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
    'СПРИНТ №' + sprintNumber,
    DARK,
    'white',
    true,
    10
  );

  sheet.setRowHeight(
    startRow,
    TITLE_ROW_HEIGHT
  );

  // Выбор спринта и дата начала
  const selectorRow = startRow + 2;

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
    10
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
    10
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
    HEADER_ROW_HEIGHT
  );

  // Разделитель
  sheet.setRowHeight(
    startRow + 3,
    SEPARATOR_ROW_HEIGHT
  );

  // Прогресс и дата окончания
  const progressRow = startRow + 4;

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
    10
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
    10
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
    HEADER_ROW_HEIGHT
  );

  // Заголовок таблицы гипотез
  const hypHeaderRow = startRow + 6;
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
      10
    );

    currentCol += width;
  });

  sheet.setRowHeight(
    hypHeaderRow,
    HYP_HEADER_ROW_HEIGHT
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
      HYP_ROW_HEIGHT
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
  const summaryGapRow = lastHypRow + 1;

  sheet.setRowHeight(
    summaryGapRow,
    SUMMARY_TOP_GAP_HEIGHT
  );

  // Итоги спринта
  const summaryHeaderRow = summaryGapRow + 1;
  const summaryValueRow = summaryHeaderRow + 1;

  summaryConfig.forEach(item => {
    mergeText(
      sheet,
      summaryHeaderRow,
      item.range.START_COL,
      summaryHeaderRow,
      item.range.END_COL,
      item.title,
      LIGHT,
      'black',
      true,
      10
    );

    mergeEmpty(
      sheet,
      summaryValueRow,
      item.range.START_COL,
      summaryValueRow,
      item.range.END_COL
    );

    setInputBorder_(
      sheet,
      summaryValueRow,
      item.range.START_COL,
      item.range.END_COL,
      BORDER
    );
  });

  sheet.setRowHeight(
    summaryHeaderRow,
    SUMMARY_HEADER_ROW_HEIGHT
  );

  sheet.setRowHeight(
    summaryValueRow,
    SUMMARY_VALUE_ROW_HEIGHT
  );

  // Пустая строка между спринтами
  sheet.setRowHeight(
    summaryValueRow + 1,
    BOTTOM_GAP_HEIGHT
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
  const GAP_HEIGHT = 8;

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

  mergeText(sheet, startRow, 1, startRow, TOTAL_COLS, title, DARK, 'white', true, 10);
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

    mergeText(sheet, r, COL_LEFT_LABEL_START, r, COL_LEFT_LABEL_END, leftItems[i], LIGHT, 'black', true, 10);
    mergeEmpty(sheet, r, COL_LEFT_VALUE_START, r, COL_LEFT_VALUE_END);
    setInputBorder_(sheet, r, COL_LEFT_VALUE_START, COL_LEFT_VALUE_END, BORDER);

    mergeText(sheet, r, COL_RIGHT_LABEL_START, r, COL_RIGHT_LABEL_END, rightItems[i], LIGHT, 'black', true, 10);
    mergeEmpty(sheet, r, COL_RIGHT_VALUE_START, r, COL_RIGHT_VALUE_END);
    setInputBorder_(sheet, r, COL_RIGHT_VALUE_START, COL_RIGHT_VALUE_END, BORDER);

    sheet.setRowHeight(r, DATA_ROW_HEIGHT);
  }

  sheet.setRowHeight(startRow + 5, GAP_HEIGHT);

  mergeText(sheet, startRow + 6, COL_PROGRESS_LABEL_START, startRow + 6, COL_PROGRESS_LABEL_END, progressTitle, LIGHT, 'black', true, 10);
  mergeEmpty(sheet, startRow + 6, COL_PROGRESS_PERCENT_START, startRow + 6, COL_PROGRESS_PERCENT_END);
  setInputBorder_(sheet, startRow + 6, COL_PROGRESS_PERCENT_START, COL_PROGRESS_PERCENT_END, BORDER);

  mergeEmpty(sheet, startRow + 6, COL_PROGRESS_BAR_START, startRow + 6, COL_PROGRESS_BAR_END);
  setInputBorder_(sheet, startRow + 6, COL_PROGRESS_BAR_START, COL_PROGRESS_BAR_END, '#000000');

  sheet.setRowHeight(startRow + 6, DATA_ROW_HEIGHT);

  return startRow + 8;
}

