/**
 * ACC KPI Framework
 * Module: Validation
 *
 * Проверяет контекст выполнения и допустимость данных.
 */

/**
 * Возвращает активный рабочий лист аккаунта.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Активный лист аккаунта.
 * @throws {Error} Если активный лист не является листом аккаунта.
 */
function getActiveAccountSheetV1_() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getActiveSheet();

  const name = sheet.getName();

  if (name === SHEETS.TEMPLATE) {
    return sheet;
  }

  if (name === SHEETS.ACCOUNTS) {
    throw new Error("Откройте лист аккаунта.");
  }

  if (!name.startsWith(SHEETS.ACCOUNT_PREFIX)) {
    throw new Error(`Лист "${name}" не является листом аккаунта.`);
  }

  return sheet;
}

function requireAdmin_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeEmail = String(
    Session.getActiveUser().getEmail() || ''
  ).trim().toLowerCase();
  const owner = ss.getOwner();
  const ownerEmail = String(
    owner && owner.getEmail ? owner.getEmail() : ''
  ).trim().toLowerCase();

  if (!activeEmail || !ownerEmail || activeEmail !== ownerEmail) {
    throw new Error(
      'Операция доступна только владельцу таблицы с подтверждённым email.'
    );
  }
}
