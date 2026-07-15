/**
 * ==========================================
 * updateSprintTotalsV1
 * ==========================================
 *
 * Назначение:
 * Пересчитывает итоги всех спринтов.
 *
 * Считает:
 * • Заполнено
 * • Выполнено
 * • В работе
 * • Отложено
 * • Прогресс
 *
 * Обновляет:
 * • Итоги спринта
 * • Процент прогресса в шапке спринта
 * • ProgressBar в шапке спринта
 */
function updateSprintTotalsV1() {
  const sheet = getActiveAccountSheetV1_();

  const COL_HYPOTHESIS = 3;
  const COL_STATUS = 75;

  const COL_TOTAL_FILLED = 5;
  const COL_TOTAL_DONE = 23;
  const COL_TOTAL_IN_PROGRESS = 41;
  const COL_TOTAL_POSTPONED = 59;
  const COL_TOTAL_PROGRESS = 77;

  const COL_HEADER_PERCENT_START = 11;
  const COL_HEADER_PERCENT_END = 14;
  const COL_HEADER_BAR_START = 15;
  const COL_HEADER_BAR_END = 36;

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
      const sprintLayout = getSprintLayoutV1_(
        monthIndex,
        sprintIndex
      );

      const hypotheses = sheet
        .getRange(
          sprintLayout.firstHypRow,
          COL_HYPOTHESIS,
          LAYOUT.SPRINT.HYPOTHESIS_COUNT,
          1
        )
        .getValues()
        .flat();

      const statuses = sheet
        .getRange(
          sprintLayout.firstHypRow,
          COL_STATUS,
          LAYOUT.SPRINT.HYPOTHESIS_COUNT,
          1
        )
        .getValues()
        .flat();

      const filled = hypotheses.filter(
        value => String(value).trim() !== ''
      ).length;

      const done = statuses.filter(
        value => value === 'Завершено'
      ).length;

      const inProgress = statuses.filter(
        value => value === 'В работе'
      ).length;

      const postponed = statuses.filter(
        value => value === 'Отложено'
      ).length;

      const progress =
        filled > 0
          ? done / filled
          : 0;

      sheet
        .getRange(
          sprintLayout.totalsRow,
          COL_TOTAL_FILLED
        )
        .setValue(filled);

      sheet
        .getRange(
          sprintLayout.totalsRow,
          COL_TOTAL_DONE
        )
        .setValue(done);

      sheet
        .getRange(
          sprintLayout.totalsRow,
          COL_TOTAL_IN_PROGRESS
        )
        .setValue(inProgress);

      sheet
        .getRange(
          sprintLayout.totalsRow,
          COL_TOTAL_POSTPONED
        )
        .setValue(postponed);

      sheet
        .getRange(
          sprintLayout.totalsRow,
          COL_TOTAL_PROGRESS
        )
        .setValue(progress)
        .setNumberFormat('0%');

      const percentRange = sheet.getRange(
        sprintLayout.progressRow,
        COL_HEADER_PERCENT_START,
        1,
        COL_HEADER_PERCENT_END -
          COL_HEADER_PERCENT_START +
          1
      );

      try {
        percentRange.breakApart();
      } catch (error) {}

      percentRange.merge();

      percentRange
        .setValue(progress)
        .setNumberFormat('0%')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setFontWeight('bold');

      const percentCell = sheet.getRange(
        sprintLayout.totalsRow,
        COL_TOTAL_PROGRESS
      );

      const barRange = sheet.getRange(
        sprintLayout.progressRow,
        COL_HEADER_BAR_START,
        1,
        COL_HEADER_BAR_END -
          COL_HEADER_BAR_START +
          1
      );

      insertProgressBarV1_(
        sheet,
        percentCell,
        barRange
      );
    }
  }
}

/**
 * ==========================================
 * updateMonthTotalsV1
 * ==========================================
 *
 * Пересчитывает итоги месяца
 * на основании итогов спринтов.
 */
function updateMonthTotalsV1() {
  const sheet = getActiveAccountSheetV1_();

  const MONTHS = 3;
  const MONTH_HEIGHT = 108;

  const FIRST_MONTH_SUMMARY_ROW = 63;
  const FIRST_MONTH_FIRST_SPRINT_TOTAL_ROW = 84;

  const SPRINTS_PER_MONTH = 5;
  const SPRINT_STEP = 15;

  const COL_TOTAL = 5; // E
  const COL_DONE = 23; // W

  const COL_LEFT_VALUE = 19; // S
  const COL_RIGHT_VALUE = 39; // AM

  const COL_PROGRESS_BAR_START = 23; // W
  const COL_PROGRESS_BAR_WIDTH = 20; // W:AP

  const PARTIAL_SUCCESS_WEIGHT = 0.5;

  for (let month = 0; month < MONTHS; month++) {
    const summaryRow = FIRST_MONTH_SUMMARY_ROW + month * MONTH_HEIGHT;
    const firstSprintTotalsRow = FIRST_MONTH_FIRST_SPRINT_TOTAL_ROW + month * MONTH_HEIGHT;

    let total = 0;
    let done = 0;

    for (let sprint = 0; sprint < SPRINTS_PER_MONTH; sprint++) {
      const totalsRow = firstSprintTotalsRow + sprint * SPRINT_STEP;

      total += Number(sheet.getRange(totalsRow, COL_TOTAL).getValue()) || 0;
      done += Number(sheet.getRange(totalsRow, COL_DONE).getValue()) || 0;
    }

    const success = Number(sheet.getRange(summaryRow, COL_RIGHT_VALUE).getValue()) || 0;
    const partial = Number(sheet.getRange(summaryRow + 1, COL_RIGHT_VALUE).getValue()) || 0;

    const conversion = done > 0
      ? (success + partial * PARTIAL_SUCCESS_WEIGHT) / done
      : 0;

    const progress = total > 0
      ? done / total
      : 0;

    sheet.getRange(summaryRow, COL_LEFT_VALUE).setValue(total);
    sheet.getRange(summaryRow + 1, COL_LEFT_VALUE).setValue(done);

    sheet
      .getRange(summaryRow + 2, COL_LEFT_VALUE)
      .setValue(conversion)
      .setNumberFormat('0%');

    const progressCell = sheet.getRange(summaryRow + 4, COL_LEFT_VALUE);

    progressCell
      .setValue(progress)
      .setNumberFormat('0%');

    const progressBarRange = sheet.getRange(
      summaryRow + 4,
      COL_PROGRESS_BAR_START,
      1,
      COL_PROGRESS_BAR_WIDTH
    );

    insertProgressBarV1_(sheet, progressCell, progressBarRange);
  }
}

/**
 * ==========================================
 * updateQuarterTotalsV1
 * ==========================================
 *
 * Пересчитывает итоги квартала
 * по итогам трёх месяцев.
 */
function updateQuarterTotalsV1() {

  const sheet = getActiveAccountSheetV1_();

  const quarterSummaryRow = LAYOUT.QUARTER.SUMMARY.ROW;

  const firstMonthSummaryRow = LAYOUT.MONTH.SUMMARY.FIRST_ROW;
  const monthHeight = LAYOUT.MONTH.HEIGHT;
  const monthCount = LAYOUT.MONTH.COUNT;

  const colLeft = 19;     // S
  const colRight = 39;    // AM

  const progressBarStart = 23; // W
  const progressBarWidth = 20; // W:AP

  const partialWeight = 0.5;

  let total = 0;
  let done = 0;
  let success = 0;
  let partial = 0;
  let failed = 0;

  for (let month = 0; month < monthCount; month++) {

    const row = firstMonthSummaryRow + month * monthHeight;

    total += Number(sheet.getRange(row, colLeft).getValue()) || 0;
    done += Number(sheet.getRange(row + 1, colLeft).getValue()) || 0;

    success += Number(sheet.getRange(row, colRight).getValue()) || 0;
    partial += Number(sheet.getRange(row + 1, colRight).getValue()) || 0;
    failed += Number(sheet.getRange(row + 2, colRight).getValue()) || 0;
  }

  const conversion =
    done > 0
      ? (success + partial * partialWeight) / done
      : 0;

  const progress =
    total > 0
      ? done / total
      : 0;

  // Левый блок

  sheet.getRange(quarterSummaryRow, colLeft).setValue(total);

  sheet.getRange(quarterSummaryRow + 1, colLeft).setValue(done);

  sheet.getRange(quarterSummaryRow + 2, colLeft)
    .setValue(conversion)
    .setNumberFormat("0.00%");

  // Правый блок

  sheet.getRange(quarterSummaryRow, colRight).setValue(success);

  sheet.getRange(quarterSummaryRow + 1, colRight).setValue(partial);

  sheet.getRange(quarterSummaryRow + 2, colRight).setValue(failed);

  // Прогресс

  const progressCell = sheet.getRange(quarterSummaryRow + 4, colLeft);

  progressCell
    .setValue(progress)
    .setNumberFormat("0.00%");

  const progressBarRange = sheet.getRange(
    quarterSummaryRow + 4,
    progressBarStart,
    1,
    progressBarWidth
  );

  insertProgressBarV1_(sheet, progressCell, progressBarRange);
}

function updateKpiMonthTotalsV1() {
  const sheet = getActiveAccountSheetV1_();

  const KPI_COUNT = 12;

  const MONTH_HEIGHT = LAYOUT.MONTH.HEIGHT;
  const MONTH_COUNT = LAYOUT.MONTH.COUNT;

  const MONTH1_KPI_FIRST_ROW = 48;
  const MONTH1_FIRST_HYP_ROW = 78;

  const SPRINT_COUNT = LAYOUT.SPRINT.COUNT;
  const SPRINT_STEP = LAYOUT.SPRINT.STEP;
  const HYP_ROWS = LAYOUT.SPRINT.HYP_ROWS;

  const COL_KPI = 1;       // A
  const COL_PLAN = 33;     // AG
  const COL_FACT = 47;     // AU
  const COL_PERCENT = 61;  // BI

  const COL_BAR_START = 65; // BM
  const COL_BAR_WIDTH = 20;

  const COL_HYP_KPI = 27;
  const COL_HYP_FACT = 48;

  for (let month = 0; month < MONTH_COUNT; month++) {
    const kpiFirstRow = MONTH1_KPI_FIRST_ROW + month * MONTH_HEIGHT;
    const firstHypRow = MONTH1_FIRST_HYP_ROW + month * MONTH_HEIGHT;

    for (let i = 0; i < KPI_COUNT; i++) {
      const kpiRow = kpiFirstRow + i;
      const kpiName = sheet.getRange(kpiRow, COL_KPI).getValue();

      if (!kpiName) {
        sheet.getRange(kpiRow, COL_FACT).clearContent();
        sheet.getRange(kpiRow, COL_PERCENT).clearContent();
        continue;
      }

      let factSum = 0;

      for (let sprint = 0; sprint < SPRINT_COUNT; sprint++) {
        const hypStartRow = firstHypRow + sprint * SPRINT_STEP;

        for (let r = hypStartRow; r < hypStartRow + HYP_ROWS; r++) {
          const hypKpi = sheet.getRange(r, COL_HYP_KPI).getValue();

          if (hypKpi === kpiName) {
            factSum += Number(sheet.getRange(r, COL_HYP_FACT).getValue()) || 0;
          }
        }
      }

      const plan = Number(sheet.getRange(kpiRow, COL_PLAN).getValue()) || 0;
      const percent = plan > 0 ? factSum / plan : "";

      sheet.getRange(kpiRow, COL_FACT).setValue(factSum);

      const percentCell = sheet.getRange(kpiRow, COL_PERCENT);

      if (percent === "") {
        percentCell.clearContent();
      } else {
        percentCell
          .setValue(percent)
          .setNumberFormat("0.00%");
      }

      const barRange = sheet.getRange(kpiRow, COL_BAR_START, 1, COL_BAR_WIDTH);
      insertProgressBarV1_(sheet, percentCell, barRange);
    }
  }
}

/**
 * ==========================================
 * updateKpiQuarterTotalsV1
 * ==========================================
 *
 * Рассчитывает % выполнения KPI квартала.
 */
function updateKpiQuarterTotalsV1() {
  const sheet = getActiveAccountSheetV1_();

  const KPI_COUNT = 12;

  const QUARTER_FIRST_ROW = 21;
  const MONTH1_FIRST_ROW = 48;

  const MONTH_HEIGHT = LAYOUT.MONTH.HEIGHT;

  const COL_KPI = 1;       // A
  const COL_PLAN = 33;     // AG
  const COL_FACT = 47;     // AU
  const COL_PERCENT = 61;  // BI

  const COL_BAR_START = 65; // BM
  const COL_BAR_WIDTH = 20;

  for (let i = 0; i < KPI_COUNT; i++) {
    const quarterRow = QUARTER_FIRST_ROW + i;
    const kpi = sheet.getRange(quarterRow, COL_KPI).getValue();

    if (!kpi) {
      sheet.getRange(quarterRow, COL_FACT).clearContent();
      sheet.getRange(quarterRow, COL_PERCENT).clearContent();
      continue;
    }

    let factSum = 0;

    for (let month = 0; month < LAYOUT.MONTH.COUNT; month++) {
      const monthRow = MONTH1_FIRST_ROW + month * MONTH_HEIGHT + i;
      factSum += Number(sheet.getRange(monthRow, COL_FACT).getValue()) || 0;
    }

    const plan = Number(sheet.getRange(quarterRow, COL_PLAN).getValue()) || 0;
    const percent = plan > 0 ? factSum / plan : "";

    sheet.getRange(quarterRow, COL_FACT).setValue(factSum);

    const percentCell = sheet.getRange(quarterRow, COL_PERCENT);

    if (percent === "") {
      percentCell.clearContent();
    } else {
      percentCell.setValue(percent).setNumberFormat("0.00%");
    }

    const barRange = sheet.getRange(quarterRow, COL_BAR_START, 1, COL_BAR_WIDTH);
    insertProgressBarV1_(sheet, percentCell, barRange);
  }
}

/**
 * ==========================================
 * recalculateAnalyticsV1
 * ==========================================
 *
 * Полный пересчёт аналитики.
 */
function recalculateAnalyticsV1() {

  // Спринты
  updateSprintTotalsV1();

  // Месяц
  updateMonthTotalsV1();

  // Квартал
  updateQuarterTotalsV1();

  // KPI
  updateKpiMonthTotalsV1();
  updateKpiQuarterTotalsV1();

}

/**
 * Копирует KPI квартала в месяцы и распределяет план.
 *
 * KPI с единицей измерения "%" — 2 знака после запятой.
 * Остальные KPI — целые значения.
 */
function calculateKpiMonthPlansV1() {
  const sheet = getActiveAccountSheetV1_();
  const kpiDict = getKpiDictionaryV1_();

  const KPI_COUNT = 12;
  const QUARTER_FIRST_ROW = 21;
  const MONTH1_FIRST_ROW = 48;

  const COL_KPI = 1;   // A
  const COL_PLAN = 33; // AG

  for (let i = 0; i < KPI_COUNT; i++) {
    const quarterRow = QUARTER_FIRST_ROW + i;

    const kpi = String(
      sheet.getRange(quarterRow, COL_KPI).getValue()
    ).trim();

    const quarterPlan = sheet
      .getRange(quarterRow, COL_PLAN)
      .getValue();

    const kpiInfo = kpiDict[kpi] || {};

    const distribution =
      kpiInfo.distribution || 'UNIFORM';

    // Процентные KPI, например ER/ERV, сохраняем до 2 знаков.
    const decimals =
      String(kpiInfo.unit).trim() === '%'
        ? 2
        : 0;

    const monthPlans =
      quarterPlan === '' || quarterPlan === null
        ? ['', '', '']
        : calculateMonthPlanV1_(
            Number(quarterPlan),
            distribution,
            decimals
          );

    for (let month = 0; month < LAYOUT.MONTH.COUNT; month++) {
      const monthRow =
        MONTH1_FIRST_ROW +
        month * LAYOUT.MONTH.HEIGHT +
        i;

      sheet.getRange(monthRow, COL_KPI).setValue(kpi);

      const planCell = sheet.getRange(monthRow, COL_PLAN);

      if (monthPlans[month] === '') {
        planCell.clearContent();
        continue;
      }

      planCell
        .setValue(monthPlans[month])
        .setNumberFormat(decimals === 2 ? '0.00' : '0');
    }
  }
}

/**
 * Пересчёт аналитики спринтов.
 */
function recalculateSprintAnalyticsV1() {

  updateSprintTotalsV1();
  updateMonthTotalsV1();
  updateQuarterTotalsV1();

}


/**
 * Пересчёт KPI.
 */
function recalculateKpiAnalyticsV1() {

  updateKpiMonthTotalsV1();
  updateKpiQuarterTotalsV1();

}

/**
 * ==========================================
 * recalculateAllAccountsKpiV1
 * ==========================================
 *
 * Пересчитывает KPI всех аккаунтов.
 */
function recalculateAllAccountsKpiV1() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();

  const accounts = ss.getSheetByName("ACCOUNTS");

  if (!accounts) {
    throw new Error('Лист "ACCOUNTS" не найден.');
  }

  const lastRow = accounts.getLastRow();

  if (lastRow < 2) return;

  const sheetNames = accounts
    .getRange(2, 3, lastRow - 1, 1)
    .getValues()
    .flat();

  let processed = 0;
  let errors = 0;

  for (const name of sheetNames) {

    if (!name) continue;

    const sheet = ss.getSheetByName(name);

    if (!sheet) {
      errors++;
      continue;
    }

    try {

      sheet.activate();

      calculateKpiMonthPlansV1();
      updateKpiMonthTotalsV1();
      updateKpiQuarterTotalsV1();

      processed++;

    } catch (e) {

      Logger.log(name + " : " + e);
      errors++;

    }

  }

  currentSheet.activate();

  SpreadsheetApp.getUi().alert(
    "Пересчёт KPI завершён.\n\n" +
    "Обработано: " + processed +
    "\nОшибок: " + errors
  );

}


/**
 * Полный пересчёт.
 */
function recalculateAnalyticsV1() {

  recalculateSprintAnalyticsV1();
  recalculateKpiAnalyticsV1();

}



/**
 * Возвращает месячный план KPI.
 *
 * @param {number} quarterValue
 * @param {string} distribution
 * @return {number[]}
 */
function calculateMonthPlanV1_(quarterValue, distribution) {

  quarterValue = Number(quarterValue) || 0;

  switch (distribution) {

    case "GROWTH10": {

      const m1 = round2_(quarterValue * 1.0 / 3.3);
      const m2 = round2_(quarterValue * 1.1 / 3.3);
      const m3 = round2_(quarterValue - m1 - m2);

      return [m1, m2, m3];
    }

    case "UNIFORM":
    default: {

      const m1 = round2_(quarterValue / 3);
      const m2 = round2_(quarterValue / 3);
      const m3 = round2_(quarterValue - m1 - m2);

      return [m1, m2, m3];
    }

  }

}

/**
 * Округляет число до заданного количества знаков.
 */
function roundTo_(value, decimals) {
  const factor = Math.pow(10, decimals);

  return Math.round(
    (Number(value) + Number.EPSILON) * factor
  ) / factor;
}

//служебные
function round2_(value) {
  return Math.round(value * 100) / 100;
}







function isHypothesisEditV1_(row, col) {
  const monthFirstHypRows = [57, 158, 259];

  const editableCols = [
    3,   // Гипотеза
    21,  // KPI
    33,  // План
    36,  // Факт
    47,  // Исполнитель
    59,  // Статус
    65   // Результат
  ];

  if (!editableCols.includes(col)) return false;

  return monthFirstHypRows.some(monthStartRow => {
    for (let sprint = 0; sprint < 5; sprint++) {
      const firstHypRow = monthStartRow + sprint * 14;
      if (row >= firstHypRow && row <= firstHypRow + 3) return true;
    }
    return false;
  });
}




