/**
 * ACC KPI Framework
 * Module: Constants
 *
 * Содержит глобальные неизменяемые значения проекта.
 *
 * ВАЖНО:
 * PASSPORT, HYP, SPRINT, MONTH и QUARTER временно сохраняются
 * для обратной совместимости. В следующих задачах их геометрия
 * будет перенесена в LAYOUT, после чего устаревшие объекты будут удалены.
 */

/* ==========================================================================
 * LEGACY LAYOUT CONSTANTS
 * Временно сохранены для обратной совместимости.
 * Не использовать в новом коде.
 * ========================================================================== */

/**
 * Устаревшие адреса полей паспорта аккаунта.
 *
 * @deprecated Использовать LAYOUT.PASSPORT и Geometry API.
 */
const PASSPORT = {
  ACCOUNT: "M4",
  SOCIAL: "M5",
  EMPLOYEE: "M6",
  YEAR: "M7",
  QUARTER: "M8",
  STATUS: "M9",
};

/**
 * Устаревшее описание строк и столбцов гипотез.
 *
 * @deprecated Использовать LAYOUT.HYPOTHESIS и Geometry API.
 */
const HYP = {
  ROWS_PER_SPRINT: 4,

  COL: {
    KPI: 27,
    PLAN: 43,
    FACT: 48,
    EXPECTED_EFFECT: 53,
    ACTUAL_EFFECT: 58,
    EMPLOYEE: 63,
    STATUS: 75,
    RESULT: 83,
  },
};

/**
 * Устаревшие параметры спринта.
 *
 * @deprecated Использовать LAYOUT.SPRINT и Geometry API.
 */
const SPRINT = {
  COUNT: 5,
  STEP_ROWS: 15,
  HEADER_HEIGHT: 32,
  HEADER_INPUT_HEIGHT: 28,
  HYPOTHESIS_HEIGHT: 48,
};

/**
 * Устаревшие параметры месяца.
 *
 * @deprecated Использовать LAYOUT.MONTH.
 */
const MONTH = {
  COUNT: 3,
};

/**
 * Устаревшие параметры квартала.
 *
 * @deprecated Использовать LAYOUT.QUARTER.
 */
const QUARTER = {
  KPI_ROWS: 11,
};

/* ==========================================================================
 * TECHNICAL CONFIGURATION
 * Техническая конфигурация проекта.
 * ========================================================================== */

/**
 * Канонические имена листов Google Sheets.
 */
const SHEETS = {
  DIRECTORY: "90_DIRECTORY",
  ACCOUNTS: "ACCOUNTS",
  TEMPLATE: "ACC_TEMPLATE",
  ACCOUNT_PREFIX: "ACC_",
};

/* ==========================================================================
 * UI CONSTANTS
 * Неизменяемые значения оформления.
 * ========================================================================== */

/**
 * Основная палитра интерфейса.
 */
const COLORS = {
  DARK: "#0B5D1E",
  LIGHT: "#EEF7EA",
  BORDER: "#D9E0D6",
  WHITE: "#FFFFFF",
  BLACK: "#000000",
};

/* ==========================================================================
 * BUSINESS CONSTANTS
 * Неизменяемые бизнес-значения.
 * ========================================================================== */

/**
 * Бизнес-коэффициенты системы.
 */
const BUSINESS = {
  PARTIAL_SUCCESS_WEIGHT: 0.5,
};

/**
 * Допустимые статусы гипотез.
 */
const STATUS = {
  DONE: "Завершено",
  IN_PROGRESS: "В работе",
  POSTPONED: "Отложено",
};

/**
 * Допустимые результаты гипотез.
 */
const RESULT = {
  SUCCESS: "Успешно",
  PARTIAL: "Частично успешно",
  FAILED: "Неуспешно",
};
