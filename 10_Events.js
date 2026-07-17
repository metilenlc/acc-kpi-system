function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  const isAccountSheet =
    sheetName === 'ACC_TEMPLATE' ||
    sheetName.startsWith('ACC_');

  if (!isAccountSheet) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  const a1 = e.range.getA1Notation();

  // Изменение квартала
  if (a1 === 'M8') {
    try {
      updateSprintDropdownsByQuarterV1(sheet);
    } catch (error) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        error.message,
        'ACC',
        5
      );
    }
    return;
  }

  // Выбор спринта
  if (col === 11 && getSprintInputRowsV1_().includes(row)) {
    updateSprintDatesOnEdit_(e);
    return;
  }

  // Изменение квартального плана KPI
  if (col === 33 && row >= 21 && row <= 32) {
    calculateKpiMonthPlansV1();
    recalculateKpiAnalyticsV1();
    return;
  }

  // Изменение гипотезы
  if (isAnalyticsEditV1_(row, col)) {
    recalculateAnalyticsV1();
  }
}

function isAnalyticsEditV1_(row, col) {
  const editableCols = [
    LAYOUT.HYPOTHESIS.COLUMNS.NAME,
    LAYOUT.HYPOTHESIS.COLUMNS.KPI,
    LAYOUT.HYPOTHESIS.COLUMNS.PLAN,
    LAYOUT.HYPOTHESIS.COLUMNS.FACT,
    LAYOUT.HYPOTHESIS.COLUMNS.STATUS,
    LAYOUT.HYPOTHESIS.COLUMNS.RESULT
  ];

  if (!editableCols.includes(col)) return false;

  for (
    let monthIndex = 0;
    monthIndex < LAYOUT.MONTH.COUNT;
    monthIndex++
  ) {
    for (
      let sprintIndex = 0;
      sprintIndex < LAYOUT.SPRINT.COUNT;
      sprintIndex++
    ) {
      const sprint = getSprintLayoutV1_(
        monthIndex,
        sprintIndex
      );

      if (
        row >= sprint.firstHypRow &&
        row <= sprint.lastHypRow
      ) {
        return true;
      }
    }
  }

  return false;
}
