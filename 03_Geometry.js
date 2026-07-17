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
 * Высота одного блока спринта.
 */
function getSprintStepV1_() {
  return (
    LAYOUT.SPRINT.HEADER_ROWS +
    LAYOUT.SPRINT.HYPOTHESIS_COUNT +
    LAYOUT.SPRINT.SUMMARY_ROWS
  );
}

/**
 * Высота одного блока месяца.
 */
function getMonthHeightV1_() {
  return (
    LAYOUT.MONTH.SPRINT_OFFSET +
    LAYOUT.SPRINT.COUNT * getSprintStepV1_() +
    LAYOUT.MONTH.TRAILING_ROWS
  );
}

//первый месяц
function getMonthFirstRowV1_(monthIndex) {
  return (
    LAYOUT.MONTH.FIRST_ROW +
    monthIndex * getMonthHeightV1_()
  );
}

//первый KPI месяца
function getMonthFirstKpiRowV1_(monthIndex) {
  return getMonthLayoutV1_(monthIndex).firstKpiDataRow;
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
  return getSprintLayoutV1_(
    monthIndex,
    sprintIndex
  ).firstHypRow;
}

//итоги спринта
function getSprintTotalsRowV1_(
  monthIndex,
  sprintIndex
) {
  return getSprintLayoutV1_(
    monthIndex,
    sprintIndex
  ).totalsRow;
}

//строка прогресс-бар
function getSprintProgressRowV1_(
  monthIndex,
  sprintIndex
) {
  return getSprintLayoutV1_(
    monthIndex,
    sprintIndex
  ).progressRow;
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
    monthIndex * getMonthHeightV1_();

  const kpiRow =
    monthRow +
    LAYOUT.MONTH.KPI_OFFSET;

  const kpiHeaderRow =
    kpiRow +
    LAYOUT.MONTH.KPI_HEADER_OFFSET_FROM_KPI_TITLE;

  const firstKpiDataRow =
    kpiRow +
    LAYOUT.MONTH.KPI_DATA_OFFSET_FROM_KPI_TITLE;

  const lastKpiDataRow =
    firstKpiDataRow +
    LAYOUT.MONTH.KPI_COUNT - 1;

  return {

    monthRow,

    titleRow: monthRow,

    kpiRow,

    kpiHeaderRow,

    firstKpiDataRow,

    lastKpiDataRow,

    kpiBottomGapRow:
      lastKpiDataRow + 1,

    summaryRow:
      monthRow +
      LAYOUT.MONTH.SUMMARY_OFFSET,

    summaryValueRow:
      monthRow +
      LAYOUT.MONTH.SUMMARY_OFFSET +
      LAYOUT.MONTH.SUMMARY_ROW_OFFSETS.FIRST_VALUE,

    summaryProgressRow:
      monthRow +
      LAYOUT.MONTH.SUMMARY_OFFSET +
      LAYOUT.MONTH.SUMMARY_ROW_OFFSETS.PROGRESS,

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

  const summaryOffsets =
    LAYOUT.SPRINT.SUMMARY_OFFSETS;

  return {

    sprintTop,

    selectorRow:
      sprintTop +
      LAYOUT.SPRINT.ROW_OFFSETS.SELECTOR,

    progressRow:
      sprintTop +
      LAYOUT.SPRINT.ROW_OFFSETS.PROGRESS,

    headerRow:
      sprintTop +
      LAYOUT.SPRINT.ROW_OFFSETS.HYPOTHESIS_HEADER,

    firstHypRow,

    lastHypRow,

    summaryGapRow:
      lastHypRow + summaryOffsets.GAP,

    totalsHeaderRow:
      lastHypRow + summaryOffsets.HEADER,

    totalsRow:
      lastHypRow + summaryOffsets.VALUES,

    bottomGapRow:
      lastHypRow + summaryOffsets.BOTTOM

  };

}
