// @ts-nocheck
/*
==================================================
01_LayoutMap.gs
==================================================
*/

const LAYOUT = {

  QUARTER: {

    KPI: {
      ROW: 20
    },

    SUMMARY: {
      ROW: 36
    }

  },

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

    SUMMARY: {
      FIRST_ROW: 63
    },

    HYPOTHESIS: {
      FIRST_ROW: 78
    },

    SPRINT: {
      FIRST_TOTAL_ROW: 84
    }

  },

  SPRINT: {

    COUNT: 5,
    HYPOTHESIS_COUNT: 4,
    HEADER_ROWS: 7,
    SUMMARY_ROWS: 4,
   
  },

  PASSPORT: {
    ACCOUNT: "M4",
    SOCIAL: "M5",
    MANAGER: "M6",
    YEAR: "M7",
    QUARTER: "M8",
    STATUS: "M9"
  }, 

  DIRECTORY: {

    KPI: {

      FIRST_ROW: 2,
      FIRST_COLUMN: 18,   // R
      COLUMN_COUNT: 8

    }
    
  },

  SPRINT_SUMMARY: {

  FILLED: {
    START_COL: 4,
    END_COL: 11
  },

  DONE: {
    START_COL: 15,
    END_COL: 22
  },

  IN_PROGRESS: {
    START_COL: 26,
    END_COL: 33
  },

  POSTPONED: {
    START_COL: 37,
    END_COL: 44
  },

  PROGRESS: {
    START_COL: 50,
    END_COL: 57
  },

  SUCCESS: {
    START_COL: 63,
    END_COL: 70
  },

  PARTIAL: {
    START_COL: 73,
    END_COL: 80
  },

  FAILED: {
    START_COL: 83,
    END_COL: 90
  }
},

};