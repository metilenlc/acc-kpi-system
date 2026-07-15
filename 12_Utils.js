function mergeText(sheet, r1, c1, r2, c2, text, bg, color, bold, size) {
  const range = sheet.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
  range.merge();
  range.setValue(text);

  if (bg) range.setBackground(bg);
  if (color) range.setFontColor(color);

  range
    .setFontWeight(bold ? 'bold' : 'normal')
    .setFontSize(size || 11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
}

function mergeEmpty(sheet, r1, c1, r2, c2) {
  sheet.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1).merge();
}

/**
 * Рамка для поля ввода.
 */
/**
 * ==========================================
 * setInputBorder_
 * ==========================================
 *
 * Назначение:
 * Добавляет внешнюю рамку для поля ввода или визуального блока.
 */
function setInputBorder_(sheet, row, colStart, colEnd, color) {

  const BORDER_COLOR = color || '#D9E0D6';

  sheet
    .getRange(
      row,
      colStart,
      1,
      colEnd - colStart + 1
    )
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      BORDER_COLOR,
      SpreadsheetApp.BorderStyle.SOLID
    );
}

/*
writeCentered_()
*/