const PROGRESS_BAR_TEMPLATE =
'=IFERROR(SPARKLINE({CELL};{"charttype"\\"bar";"max"\\100%;"color1"\\if({CELL}>=0,8;"#8CC474";if({CELL}>=0,7;"#EBE964";if({CELL}>=0,4;"#FAB012";"#F56B5D")))}))';

function insertProgressBarV1_(sheet, percentCell, barRange) {
  const formula = PROGRESS_BAR_TEMPLATE.replaceAll(
    '{CELL}',
    percentCell.getA1Notation()
  );

  try {
    barRange.breakApart();
  } catch (e) {}

  barRange.merge();

  barRange
    .setFormula(formula)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  barRange.setBorder(
    true,   // верх
    true,   // слева
    true,   // снизу
    true,   // справа
    false,  // вертикальные внутренние
    false,  // горизонтальные внутренние
    "#000000",
    SpreadsheetApp.BorderStyle.SOLID);

    // белый фон
    barRange.setBackground("#FFFFFF");
}

function clearOldSprintProgressBarsV1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ACC_TEMPLATE');

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
      const progressRow =
        getSprintLayoutV1_(monthIndex, sprintIndex).progressRow;
      const bar = LAYOUT.SPRINT.PROGRESS.BAR;
      const barRange = sheet.getRange(
        progressRow,
        bar.START_COLUMN,
        1,
        bar.END_COLUMN - bar.START_COLUMN + 1
      );

      barRange.breakApart();
      barRange.clearContent();
    }
  }
}



/*
setupSprintProgressBarsV1()
setupKpiProgressBarsV1()
setupMonthProgressBarsV1()
setupQuarterProgressBarsV1()
*/
