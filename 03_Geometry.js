let KPI_DICTIONARY_V1 = null;

/**
 * Возвращает активный рабочий лист аккаунта.
 */
function getActiveAccountSheetV1_() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getActiveSheet();

  const name = sheet.getName();

  if (name === "ACC_TEMPLATE") {
    return sheet;
  }

  if (name === "ACCOUNTS") {
    throw new Error("Откройте лист аккаунта.");
  }

  if (!name.startsWith("ACC_")) {
    throw new Error(`Лист "${name}" не является листом аккаунта.`);
  }

  return sheet;
}

/**
 * Открывает лист справочника KPI.
 */
function openKpiDirectoryV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('KPI');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Лист "KPI" не найден.');
    return;
  }

  sheet.activate();
  sheet.setActiveSelection('A1');
}

/**
 * Загружает справочник KPI из листа 90_DIRECTORY.
 */
function getKpiDictionaryV1_() {
  if (KPI_DICTIONARY_V1) {
    return KPI_DICTIONARY_V1;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DIRECTORY);

  if (!sheet) {
    throw new Error(`Лист "${SHEETS.DIRECTORY}" не найден.`);
  }

  const cfg = LAYOUT.DIRECTORY.KPI;
  const lastRow = sheet.getLastRow();

  if (lastRow < cfg.FIRST_ROW) {
    return {};
  }

  const rowCount = lastRow - cfg.FIRST_ROW + 1;

  const values = sheet
    .getRange(
      cfg.FIRST_ROW,
      cfg.FIRST_COLUMN,
      rowCount,
      cfg.COLUMN_COUNT
    )
    .getValues();

  const dict = {};

  values.forEach(row => {
    const kpi = String(row[0] || '').trim();

    if (!kpi) return;

    dict[kpi] = {
      unit: row[1],
      aggregation: row[2],
      distribution: row[3] || 'UNIFORM',

      k1: Number(row[4]) || 1,
      k2: Number(row[5]) || 1,
      k3: Number(row[6]) || 1
    };
  });

  KPI_DICTIONARY_V1 = dict;

  return dict;
}

/**
 * Высота одного блока спринта.
 */
function getSprintStepV1_() {
  return (
    LAYOUT.SPRINT.HEADER_ROWS +
    LAYOUT.SPRINT.HYPOTHESIS_COUNT +
    LAYOUT.SPRINT.SUMMARY_ROWS
  );
}

//первый месяц
function getMonthFirstRowV1_(monthIndex) {
  return (
    LAYOUT.MONTH.FIRST_ROW +
    monthIndex * LAYOUT.MONTH.HEIGHT
  );
}

//первый KPI месяца
function getMonthFirstKpiRowV1_(monthIndex) {
  return (
    getMonthFirstRowV1_(monthIndex) +
    LAYOUT.MONTH.KPI_OFFSET
  );
}

//первый спринт месяца
function getMonthFirstSprintRowV1_(monthIndex) {
  return (
    getMonthFirstRowV1_(monthIndex) +
    LAYOUT.MONTH.SPRINT_OFFSET
  );
}

//первая гипотеза спринта
function getSprintFirstHypRowV1_(
  monthIndex,
  sprintIndex
) {
  return (
    getMonthFirstSprintRowV1_(monthIndex) +
    sprintIndex * getSprintStepV1_() +
    LAYOUT.SPRINT.HEADER_ROWS
  );
}

//итоги спринта
function getSprintTotalsRowV1_(
  monthIndex,
  sprintIndex
) {
  return (
    getSprintFirstHypRowV1_(
      monthIndex,
      sprintIndex
    ) +
    LAYOUT.SPRINT.HYPOTHESIS_COUNT +
    2
  );
}

//строка прогресс-бар
function getSprintProgressRowV1_(
  monthIndex,
  sprintIndex
) {
  return (
    getSprintFirstHypRowV1_(
      monthIndex,
      sprintIndex
    ) - 3
  );
}

/**
 * Геометрия месяца.
 *
 * monthIndex:
 * 0 = Месяц 1
 * 1 = Месяц 2
 * 2 = Месяц 3
 */
function getMonthLayoutV1_(monthIndex) {

  const monthRow =
    LAYOUT.MONTH.FIRST_ROW +
    monthIndex * LAYOUT.MONTH.HEIGHT;

  return {

    monthRow,

    titleRow: monthRow,

    kpiRow:
      monthRow +
      LAYOUT.MONTH.KPI_OFFSET,

    summaryRow:
      monthRow +
      LAYOUT.MONTH.SUMMARY_OFFSET,

    firstSprintRow:
      monthRow +
      LAYOUT.MONTH.SPRINT_OFFSET

  };

}

/**
 * Геометрия спринта.
 *
 * monthIndex : 0..2
 * sprintIndex: 0..4
 */
function getSprintLayoutV1_(monthIndex, sprintIndex) {

  const month =
    getMonthLayoutV1_(monthIndex);

  const sprintTop =
    month.firstSprintRow +
    sprintIndex * getSprintStepV1_();

  const firstHypRow =
    sprintTop +
    LAYOUT.SPRINT.HEADER_ROWS;

  const lastHypRow =
    firstHypRow +
    LAYOUT.SPRINT.HYPOTHESIS_COUNT - 1;

  return {

    sprintTop,

    selectorRow: sprintTop + 2,

    progressRow: sprintTop + 4,

    headerRow: sprintTop + 6,

    firstHypRow,

    lastHypRow,

    summaryGapRow:
      lastHypRow + 1,

    totalsHeaderRow:
      lastHypRow + 2,

    totalsRow:
      lastHypRow + 3,

    bottomGapRow:
      lastHypRow + 4

  };

}