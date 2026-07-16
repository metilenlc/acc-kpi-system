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

  const progressRows = [
    54, 68, 82, 96, 110,
    155, 169, 183, 197, 211,
    256, 270, 284, 298, 312
  ];

  progressRows.forEach(r => {
    sheet.getRange(r, 49, 1, 60).breakApart();     // AW:DD
    sheet.getRange(r, 49, 1, 60).clearContent();   // AW:DD
  });
}



/*
setupSprintProgressBarsV1()
setupKpiProgressBarsV1()
setupMonthProgressBarsV1()
setupQuarterProgressBarsV1()
*/