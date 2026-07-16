// @ts-nocheck

/**
 * ACC KPI Framework
 * Module: Layout Map
 *
 * Единственный источник декларативного описания шаблона Google Sheets.
 *
 * Содержит:
 * - базовые строки и столбцы;
 * - размеры блоков;
 * - смещения;
 * - количество повторяющихся элементов;
 * - структуру паспорта, KPI, месяцев, спринтов и гипотез.
 *
 * Не содержит:
 * - SpreadsheetApp;
 * - чтение или запись данных;
 * - бизнес-расчёты;
 * - вычисление координат конкретного месяца или спринта.
 */

const LAYOUT = {
  /**
   * Общие параметры рабочего листа аккаунта.
   */
  SHEET: {
    FIRST_ROW: 1,
    FIRST_COLUMN: 1,
    LAST_COLUMN: 90,
    COLUMN_COUNT: 90,
  },

  /**
   * Паспорт аккаунта.
   */
  PASSPORT: {
    COLUMN: 13,

    FIRST_ROW: 4,
    LAST_ROW: 9,

    ACCOUNT: "M4",
    SOCIAL: "M5",
    MANAGER: "M6",
    YEAR: "M7",
    QUARTER: "M8",
    STATUS: "M9",

    ROWS: {
      ACCOUNT: 4,
      SOCIAL: 5,
      MANAGER: 6,
      YEAR: 7,
      QUARTER: 8,
      STATUS: 9,
    },
  },

  /**
   * Квартальный блок.
   */
  QUARTER: {
    KPI: {
      ROW: 20,
    },

    KPI_COUNT: 12,

    SUMMARY: {
      ROW: 36,
    },
  },

  /**
   * Повторяющийся блок месяца.
   */
  MONTH: {
    COUNT: 3,

    FIRST_ROW: 43,
    HEIGHT: 108,

    TITLE_HEIGHT: 1,

    KPI_OFFSET: 5,
    KPI_HEIGHT: 12,

    SUMMARY_OFFSET: 17,
    SUMMARY_HEIGHT: 6,

    SPRINT_OFFSET: 35,

    /**
     * Временные ключи совместимости.
     *
     * Используются текущими функциями и будут удалены после перехода
     * всех модулей на Geometry API.
     */
    SUMMARY: {
      FIRST_ROW: 63,
    },

    HYPOTHESIS: {
      FIRST_ROW: 78,
    },

    SPRINT: {
      FIRST_TOTAL_ROW: 84,
    },
  },

  /**
   * Повторяющийся блок спринта.
   */
  SPRINT: {
    COUNT: 5,
    HYPOTHESIS_COUNT: 4,

    HEADER_ROWS: 7,
    SUMMARY_ROWS: 4,

    ROW_OFFSETS: {
      TITLE: 0,
      SELECTOR: 2,
      SEPARATOR: 3,
      PROGRESS: 4,
      HYPOTHESIS_HEADER: 6,
      FIRST_HYPOTHESIS: 7,
    },

    ROW_HEIGHTS: {
      TITLE: 32,
      SELECTOR: 28,
      SEPARATOR: 6,
      PROGRESS: 28,
      HYPOTHESIS_HEADER: 32,
      HYPOTHESIS: 48,
      SUMMARY_GAP: 6,
      SUMMARY_HEADER: 26,
      SUMMARY_VALUE: 28,
      BOTTOM_GAP: 28,
    },

    SELECTOR: {
      LABEL: {
        START_COLUMN: 1,
        END_COLUMN: 10,
      },

      VALUE: {
        START_COLUMN: 11,
        END_COLUMN: 36,
      },
    },

    PROGRESS: {
      LABEL: {
        START_COLUMN: 1,
        END_COLUMN: 10,
      },

      PERCENT: {
        START_COLUMN: 11,
        END_COLUMN: 14,
      },

      BAR: {
        START_COLUMN: 15,
        END_COLUMN: 36,
      },
    },

    DATE: {
      LABEL: {
        START_COLUMN: 49,
        END_COLUMN: 57,
      },

      VALUE: {
        START_COLUMN: 58,
        END_COLUMN: 72,
      },
    },
  },

  /**
   * Таблица гипотез одного спринта.
   */
  HYPOTHESIS: {
    COLUMN_COUNT: 90,

    COLUMNS: {
      CHECK: 1,
      NAME: 3,
      KPI: 27,
      PLAN: 43,
      FACT: 48,
      EXPECTED_EFFECT: 53,
      ACTUAL_EFFECT: 58,
      MANAGER: 63,
      STATUS: 75,
      RESULT: 83,
    },

    FIELDS: [
      {
        KEY: "CHECK",
        TITLE: "✓",
        START_COLUMN: 1,
        END_COLUMN: 2,
      },
      {
        KEY: "NAME",
        TITLE: "Гипотеза",
        START_COLUMN: 3,
        END_COLUMN: 26,
      },
      {
        KEY: "KPI",
        TITLE: "KPI",
        START_COLUMN: 27,
        END_COLUMN: 42,
      },
      {
        KEY: "PLAN",
        TITLE: "План",
        START_COLUMN: 43,
        END_COLUMN: 47,
      },
      {
        KEY: "FACT",
        TITLE: "Факт",
        START_COLUMN: 48,
        END_COLUMN: 52,
      },
      {
        KEY: "EXPECTED_EFFECT",
        TITLE: "Ожид.",
        START_COLUMN: 53,
        END_COLUMN: 57,
      },
      {
        KEY: "ACTUAL_EFFECT",
        TITLE: "Факт.эфф.",
        START_COLUMN: 58,
        END_COLUMN: 62,
      },
      {
        KEY: "MANAGER",
        TITLE: "Исполнитель",
        START_COLUMN: 63,
        END_COLUMN: 74,
      },
      {
        KEY: "STATUS",
        TITLE: "Статус",
        START_COLUMN: 75,
        END_COLUMN: 82,
      },
      {
        KEY: "RESULT",
        TITLE: "Результат",
        START_COLUMN: 83,
        END_COLUMN: 90,
      },
    ],
  },

  /**
   * Сводка спринта.
   */
  SPRINT_SUMMARY: {
    FILLED: {
      TITLE: "Заполнено",
      START_COLUMN: 4,
      END_COLUMN: 11,
    },

    DONE: {
      TITLE: "Выполнено",
      START_COLUMN: 15,
      END_COLUMN: 22,
    },

    IN_PROGRESS: {
      TITLE: "В работе",
      START_COLUMN: 26,
      END_COLUMN: 33,
    },

    POSTPONED: {
      TITLE: "Отложено",
      START_COLUMN: 37,
      END_COLUMN: 44,
    },

    PROGRESS: {
      TITLE: "Прогресс",
      START_COLUMN: 50,
      END_COLUMN: 57,
    },

    SUCCESS: {
      TITLE: "Успешно",
      START_COLUMN: 63,
      END_COLUMN: 70,
    },

    PARTIAL: {
      TITLE: "Частично успешно",
      START_COLUMN: 73,
      END_COLUMN: 80,
    },

    FAILED: {
      TITLE: "Неуспешно",
      START_COLUMN: 83,
      END_COLUMN: 90,
    },
  },

  /**
   * Справочники.
   */
  DIRECTORY: {
    KPI: {
      FIRST_ROW: 2,
      FIRST_COLUMN: 18,
      COLUMN_COUNT: 7,
    },
  },
};
