/*
Выпадающие списки
*/

/**
 * Обновляет списки спринтов на переданном или активном листе
 * в соответствии с кварталом в M8.
 */
function updateSprintDropdownsByQuarterV1(targetSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = targetSheet || getActiveAccountSheetV1_();

  const quarter = String(sheet.getRange('M8').getValue()).trim();

  const rangeNameByQuarter = {
    Q1: 'Quarter_1',
    Q2: 'Quarter_2',
    Q3: 'Quarter_3',
    Q4: 'Quarter_4'
  };

  const rangeName = rangeNameByQuarter[quarter];

  if (!rangeName) {
    throw new Error(
      `На листе "${sheet.getName()}" не выбран корректный квартал в M8.`
    );
  }

  const sourceRange = ss.getRangeByName(rangeName);

  if (!sourceRange) {
    throw new Error(`Именованный диапазон "${rangeName}" не найден.`);
  }

  const ruleSprint = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .build();

  const sprintInputRows = getSprintInputRowsV1_();

  sprintInputRows.forEach(row => {
    sheet.getRange(row, 11).setDataValidation(ruleSprint); // K
  });
}

/**
 * Подключает списки спринтов ко всем рабочим листам аккаунтов.
 */
function updateSprintDropdownsForAllAccountsV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let processed = 0;
  let skipped = 0;
  const errors = [];

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();

    if (!name.startsWith('ACC_') || name === 'ACC_TEMPLATE') {
      return;
    }

    try {
      const quarter = String(sheet.getRange('M8').getValue()).trim();

      if (!quarter) {
        skipped++;
        return;
      }

      updateSprintDropdownsByQuarterV1(sheet);
      processed++;

    } catch (error) {
      errors.push(`${name}: ${error.message}`);
    }
  });

  SpreadsheetApp.getUi().alert(
    'Списки спринтов обновлены.\n\n' +
    `Обработано: ${processed}\n` +
    `Пропущено без квартала: ${skipped}\n` +
    `Ошибок: ${errors.length}` +
    (errors.length ? `\n\n${errors.slice(0, 10).join('\n')}` : '')
  );
}

//
function setupDropdownsV1(targetSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = targetSheet || getActiveAccountSheetV1_();

  const KPI_COUNT = 12;

  const COL_KPI = 1;
  const COL_SPRINT_SELECTOR = 11;
  const COL_HYP_KPI = 27;
  const COL_HYP_EMPLOYEE = 63;
  const COL_HYP_STATUS = 75;
  const COL_HYP_RESULT = 83;

  const ruleAccounts = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      ss.getRangeByName('Accounts'),
      true
    )
    .build();

  const ruleSocials = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      ss.getRangeByName('Socials'),
      true
    )
    .build();

  const ruleEmployees = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      ss.getRangeByName('Employees'),
      true
    )
    .build();

  const ruleQuarter = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      ss.getRangeByName('Quarter_Select'),
      true
    )
    .build();

  const ruleStatus = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      ss.getRangeByName('Active_Status'),
      true
    )
    .build();

  const ruleKpi = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      ss.getRangeByName('KPI_List'),
      true
    )
    .build();

  const ruleSprintStatus = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      ss.getRangeByName('Spr_Status'),
      true
    )
    .build();

  const ruleHypResult = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      ss.getRangeByName('Hyp_Result'),
      true
    )
    .build();

  // Паспорт
  sheet.getRange('M4').setDataValidation(ruleAccounts);
  sheet.getRange('M5').setDataValidation(ruleSocials);
  sheet.getRange('M6').setDataValidation(ruleEmployees);
  sheet.getRange('M8').setDataValidation(ruleQuarter);
  sheet.getRange('M9').setDataValidation(ruleStatus);

  // KPI квартала
  sheet
    .getRange(
      21,
      COL_KPI,
      KPI_COUNT,
      1
    )
    .setDataValidation(ruleKpi);

  // KPI месяцев и гипотезы
  for (
    let monthIndex = 0;
    monthIndex < LAYOUT.MONTH.COUNT;
    monthIndex++
  ) {
    const monthLayout = getMonthLayoutV1_(monthIndex);

    sheet
      .getRange(
        monthLayout.kpiRow,
        COL_KPI,
        KPI_COUNT,
        1
      )
      .setDataValidation(ruleKpi);

    for (
      let sprintIndex = 0;
      sprintIndex < LAYOUT.SPRINT.COUNT;
      sprintIndex++
    ) {
      const sprintLayout = getSprintLayoutV1_(
        monthIndex,
        sprintIndex
      );

      sheet
        .getRange(
          sprintLayout.firstHypRow,
          COL_HYP_KPI,
          LAYOUT.SPRINT.HYPOTHESIS_COUNT,
          1
        )
        .setDataValidation(ruleKpi);

      sheet
        .getRange(
          sprintLayout.firstHypRow,
          COL_HYP_EMPLOYEE,
          LAYOUT.SPRINT.HYPOTHESIS_COUNT,
          1
        )
        .setDataValidation(ruleEmployees);

      sheet
        .getRange(
          sprintLayout.firstHypRow,
          COL_HYP_STATUS,
          LAYOUT.SPRINT.HYPOTHESIS_COUNT,
          1
        )
        .setDataValidation(ruleSprintStatus);

      sheet
        .getRange(
          sprintLayout.firstHypRow,
          COL_HYP_RESULT,
          LAYOUT.SPRINT.HYPOTHESIS_COUNT,
          1
        )
        .setDataValidation(ruleHypResult);

      sheet
        .getRange(
          sprintLayout.selectorRow,
          COL_SPRINT_SELECTOR
        )
        .clearDataValidations();
    }
  }

  const quarter = String(
    sheet.getRange('M8').getValue()
  ).trim();

  if (['Q1', 'Q2', 'Q3', 'Q4'].includes(quarter)) {
    updateSprintDropdownsByQuarterV1(sheet);
  }

  SpreadsheetApp.flush();

  ss.toast(
    `Списки подключены: ${sheet.getName()}`,
    'ACC',
    3
  );
}

//удаление списков со всего листа
function clearDropdownsOnActiveSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  sheet.getDataRange().clearDataValidations();

  SpreadsheetApp.getUi().alert(
    'Выпадающие списки удалены на листе: ' + sheet.getName()
  );
}

//список спринтов
