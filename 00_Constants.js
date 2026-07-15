/*
==================================================
00_Constants.gs
==================================================
*/

//////////////////////////////////////////////////
// PASSPORT
//////////////////////////////////////////////////

const PASSPORT = {

  ACCOUNT: 'M4',

  SOCIAL: 'M5',

  EMPLOYEE: 'M6',

  YEAR: 'M7',

  QUARTER: 'M8',

  STATUS: 'M9'

};

//////////////////////////////////////////////////
// HYPOTHESIS
//////////////////////////////////////////////////

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

    RESULT: 83

  }

};

//////////////////////////////////////////////////
// SPRINT
//////////////////////////////////////////////////

const SPRINT = {

  COUNT: 5,

  STEP_ROWS: 15,

  HEADER_HEIGHT: 32,

  HEADER_INPUT_HEIGHT: 28,

  HYPOTHESIS_HEIGHT: 48

};

//////////////////////////////////////////////////
// MONTH
//////////////////////////////////////////////////

const MONTH = {

  COUNT: 3

};

//////////////////////////////////////////////////
// QUARTER
//////////////////////////////////////////////////

const QUARTER = {

  KPI_ROWS: 11

};

//////////////////////////////////////////////////
// COLORS
//////////////////////////////////////////////////

const COLORS = {

  DARK: '#0B5D1E',

  LIGHT: '#EEF7EA',

  BORDER: '#D9E0D6',

  WHITE: '#FFFFFF',

  BLACK: '#000000'

};

//////////////////////////////////////////////////
// BUSINESS RULES
//////////////////////////////////////////////////

const BUSINESS = {

  PARTIAL_SUCCESS_WEIGHT: 0.5

};

//////////////////////////////////////////////////
// SHEETS NAME
//////////////////////////////////////////////////

const SHEETS = {
  DIRECTORY: '90_DIRECTORY',
  ACCOUNTS: 'ACCOUNTS',
  TEMPLATE: 'ACC_TEMPLATE'
};

const STATUS = {
  DONE: 'Завершено',
  IN_PROGRESS: 'В работе',
  POSTPONED: 'Отложено'
};

const RESULT = {
  SUCCESS: 'Успешно',
  PARTIAL: 'Частично успешно',
  FAILED: 'Неуспешно'
};