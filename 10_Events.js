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
  const hypStartRows = [78, 186, 294];
  const sprintStep = 15;
  const sprintCount = 5;
  const hypRows = 4;

  const editableCols = [
    3,   // Гипотеза
    27,  // KPI
    43,  // План KPI
    48,  // Факт KPI
    75,  // Статус
    83   // Результат
  ];

  if (!editableCols.includes(col)) return false;

  return hypStartRows.some(monthStartRow => {
    for (let sprint = 0; sprint < sprintCount; sprint++) {
      const startRow = monthStartRow + sprint * sprintStep;
      const endRow = startRow + hypRows - 1;

      if (row >= startRow && row <= endRow) return true;
    }

    return false;
  });
}