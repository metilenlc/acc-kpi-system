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
 * • Успешно
 * • Частично успешно
 * • Неуспешно
 *
 * Обновляет:
 * • Итоги спринта
 * • Процент прогресса в шапке спринта
 * • ProgressBar в шапке спринта
 */
function updateSprintTotalsV1() {
  const sheet = getActiveAccountSheetV1_();

  const COL_HYPOTHESIS =
    LAYOUT.HYPOTHESIS.COLUMNS.NAME;
  const COL_STATUS =
    LAYOUT.HYPOTHESIS.COLUMNS.STATUS;
  const COL_RESULT =
    LAYOUT.HYPOTHESIS.COLUMNS.RESULT;

  const COL_TOTAL_FILLED =
    LAYOUT.SPRINT_SUMMARY.FILLED.START_COLUMN;
  const COL_TOTAL_DONE =
    LAYOUT.SPRINT_SUMMARY.DONE.START_COLUMN;
  const COL_TOTAL_IN_PROGRESS =
    LAYOUT.SPRINT_SUMMARY.IN_PROGRESS.START_COLUMN;
  const COL_TOTAL_POSTPONED =
    LAYOUT.SPRINT_SUMMARY.POSTPONED.START_COLUMN;
  const COL_TOTAL_PROGRESS =
    LAYOUT.SPRINT_SUMMARY.PROGRESS.START_COLUMN;
  const COL_TOTAL_SUCCESS =
    LAYOUT.SPRINT_SUMMARY.SUCCESS.START_COLUMN;
  const COL_TOTAL_PARTIAL =
    LAYOUT.SPRINT_SUMMARY.PARTIAL.START_COLUMN;
  const COL_TOTAL_FAILED =
    LAYOUT.SPRINT_SUMMARY.FAILED.START_COLUMN;

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

      const results = sheet
        .getRange(
          sprintLayout.firstHypRow,
          COL_RESULT,
          sprintLayout.lastHypRow -
            sprintLayout.firstHypRow +
            1,
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
        value => value === STATUS.POSTPONED
      ).length;

      const success = results.filter(
        value => value === RESULT.SUCCESS
      ).length;

      const partial = results.filter(
        value => value === RESULT.PARTIAL
      ).length;

      const failed = results.filter(
        value => value === RESULT.FAILED
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

      sheet
        .getRange(sprintLayout.totalsRow, COL_TOTAL_SUCCESS)
        .setValue(success);

      sheet
        .getRange(sprintLayout.totalsRow, COL_TOTAL_PARTIAL)
        .setValue(partial);

      sheet
        .getRange(sprintLayout.totalsRow, COL_TOTAL_FAILED)
        .setValue(failed);

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

  const COL_TOTAL =
    LAYOUT.SPRINT_SUMMARY.FILLED.START_COLUMN;
  const COL_DONE =
    LAYOUT.SPRINT_SUMMARY.DONE.START_COLUMN;
  const COL_SUCCESS =
    LAYOUT.SPRINT_SUMMARY.SUCCESS.START_COLUMN;
  const COL_PARTIAL =
    LAYOUT.SPRINT_SUMMARY.PARTIAL.START_COLUMN;
  const COL_FAILED =
    LAYOUT.SPRINT_SUMMARY.FAILED.START_COLUMN;

  const COL_LEFT_VALUE = 19; // S
  const COL_RIGHT_VALUE = 39; // AM

  const COL_PROGRESS_BAR_START = 23; // W
  const COL_PROGRESS_BAR_WIDTH = 20; // W:AP

  for (
    let month = 0;
    month < LAYOUT.MONTH.COUNT;
    month++
  ) {
    const monthLayout = getMonthLayoutV1_(month);
    const summaryRow = monthLayout.summaryValueRow;

    let total = 0;
    let done = 0;
    let success = 0;
    let partial = 0;
    let failed = 0;

    for (
      let sprint = 0;
      sprint < LAYOUT.SPRINT.COUNT;
      sprint++
    ) {
      const totalsRow =
        getSprintLayoutV1_(month, sprint).totalsRow;

      total += Number(sheet.getRange(totalsRow, COL_TOTAL).getValue()) || 0;
      done += Number(sheet.getRange(totalsRow, COL_DONE).getValue()) || 0;
      success += Number(sheet.getRange(totalsRow, COL_SUCCESS).getValue()) || 0;
      partial += Number(sheet.getRange(totalsRow, COL_PARTIAL).getValue()) || 0;
      failed += Number(sheet.getRange(totalsRow, COL_FAILED).getValue()) || 0;
    }

    const conversion = done > 0
      ? (success + partial * BUSINESS.PARTIAL_SUCCESS_WEIGHT) / done
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

    sheet.getRange(summaryRow, COL_RIGHT_VALUE).setValue(success);
    sheet.getRange(summaryRow + 1, COL_RIGHT_VALUE).setValue(partial);
    sheet.getRange(summaryRow + 2, COL_RIGHT_VALUE).setValue(failed);

    const progressCell = sheet.getRange(
      monthLayout.summaryProgressRow,
      COL_LEFT_VALUE
    );

    progressCell
      .setValue(progress)
      .setNumberFormat('0%');

    const progressBarRange = sheet.getRange(
      monthLayout.summaryProgressRow,
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

    const row = getMonthLayoutV1_(month).summaryValueRow;

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

  const KPI_COUNT = LAYOUT.MONTH.KPI_COUNT;

  const MONTH_COUNT = LAYOUT.MONTH.COUNT;

  const SPRINT_COUNT = LAYOUT.SPRINT.COUNT;

  const COL_KPI = 1;       // A
  const COL_PLAN = 33;     // AG
  const COL_FACT = 47;     // AU
  const COL_PERCENT = 61;  // BI

  const COL_BAR_START = 65; // BM
  const COL_BAR_WIDTH = 20;

  const COL_HYP_KPI =
    LAYOUT.HYPOTHESIS.COLUMNS.KPI;
  const COL_HYP_FACT =
    LAYOUT.HYPOTHESIS.COLUMNS.FACT;

  for (let month = 0; month < MONTH_COUNT; month++) {
    const kpiFirstRow = getMonthLayoutV1_(month).firstKpiDataRow;

    for (let i = 0; i < KPI_COUNT; i++) {
      const kpiRow = kpiFirstRow + i;
      const kpiName = String(
        sheet.getRange(kpiRow, COL_KPI).getValue()
      ).trim();

      if (!kpiName) {
        sheet.getRange(kpiRow, COL_FACT).clearContent();
        sheet.getRange(kpiRow, COL_PERCENT).clearContent();
        continue;
      }

      let factSum = 0;

      for (let sprint = 0; sprint < SPRINT_COUNT; sprint++) {
        const sprintLayout = getSprintLayoutV1_(month, sprint);

        for (
          let r = sprintLayout.firstHypRow;
          r <= sprintLayout.lastHypRow;
          r++
        ) {
          const hypKpi = String(
            sheet.getRange(r, COL_HYP_KPI).getValue()
          ).trim();

          if (hypKpi === kpiName) {
            factSum += Number(sheet.getRange(r, COL_HYP_FACT).getValue()) || 0;
          }
        }
      }

      const plan = Number(sheet.getRange(kpiRow, COL_PLAN).getValue()) || 0;
      const percent = plan > 0 ? factSum / plan : 0;

      sheet.getRange(kpiRow, COL_FACT).setValue(factSum);

      const percentCell = sheet.getRange(kpiRow, COL_PERCENT);

      percentCell
        .setValue(percent)
        .setNumberFormat("0.00%");

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

  const KPI_COUNT = LAYOUT.MONTH.KPI_COUNT;

  const QUARTER_FIRST_ROW = 21;
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
      const monthRow = getMonthLayoutV1_(month).firstKpiDataRow + i;
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
  distributeKpiPlansForSheetV1_(sheet);
}

/**
 * Distributes quarterly KPI plans on an explicitly supplied account sheet.
 * Does not activate the sheet or depend on the current UI selection.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function distributeKpiPlansForSheetV1_(sheet) {
  if (!sheet) throw new Error('Не передан лист аккаунта для распределения KPI.');
  const kpiDict = getKpiDictionaryV1_();

  const KPI_COUNT = LAYOUT.MONTH.KPI_COUNT;
  const QUARTER_FIRST_ROW = LAYOUT.QUARTER.KPI.ROW + 1;
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
        getMonthLayoutV1_(month).firstKpiDataRow +
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
 * @param {number} decimals
 * @return {number[]}
 */
function calculateMonthPlanV1_(quarterValue, distribution, decimals) {

  quarterValue = Number(quarterValue) || 0;
  const roundValue = decimals === 2
    ? value => roundTo_(value, 2)
    : value => Math.round(value);

  switch (distribution) {

    case "GROWTH10": {
      const weight1 = 1;
      const weight2 = 1.1;
      const weight3 = 1.21;
      const totalWeight = weight1 + weight2 + weight3;

      const m1 = roundValue(quarterValue * weight1 / totalWeight);
      const m2 = roundValue(quarterValue * weight2 / totalWeight);
      const m3 = decimals === 2
        ? roundTo_(quarterValue - m1 - m2, 2)
        : quarterValue - m1 - m2;

      return [m1, m2, m3];
    }

    case "UNIFORM":
    default: {

      const m1 = roundValue(quarterValue / 3);
      const m2 = roundValue(quarterValue / 3);
      const m3 = decimals === 2
        ? roundTo_(quarterValue - m1 - m2, 2)
        : quarterValue - m1 - m2;

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
  const editableCols = [
    LAYOUT.HYPOTHESIS.COLUMNS.NAME,
    LAYOUT.HYPOTHESIS.COLUMNS.KPI,
    LAYOUT.HYPOTHESIS.COLUMNS.PLAN,
    LAYOUT.HYPOTHESIS.COLUMNS.FACT,
    LAYOUT.HYPOTHESIS.COLUMNS.MANAGER,
    LAYOUT.HYPOTHESIS.COLUMNS.STATUS,
    LAYOUT.HYPOTHESIS.COLUMNS.RESULT
  ];

  if (!editableCols.includes(col)) return false;

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
      const sprint = getSprintLayoutV1_(
        monthIndex,
        sprintIndex
      );

      if (
        row >= sprint.firstHypRow &&
        row <= sprint.lastHypRow
      ) {
        return true;
      }
    }
  }

  return false;
}
