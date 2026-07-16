/**
 * ACC KPI Framework
 * Module: Directory
 *
 * Предоставляет доступ к справочникам и кэширует их данные.
 */

let KPI_DICTIONARY_V1 = null;

/**
 * Загружает и кэширует справочник KPI из канонического листа справочников.
 *
 * @returns {Object} Словарь параметров KPI по названию KPI.
 * @throws {Error} Если канонический лист справочников не найден.
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
