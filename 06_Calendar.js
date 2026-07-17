function buildMonthlySprintCalendar_2025_2030() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = '95_CALENDAR';

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  sheet.clear();
  sheet.clearFormats();

  const startYear = 2025;
  const endYear = 2030;

  const headers = [
    'Year',
    'Quarter',
    'Quarter_Name',
    'Month_Number',
    'Month_Name',
    'Sprint_ID',
    'Sprint_Number',
    'Sprint_Name',
    'Sprint_Start',
    'Sprint_End',
    'Days_Count'
  ];

  const monthNames = [
    'Январь',
    'Февраль',
    'Март',
    'Апрель',
    'Май',
    'Июнь',
    'Июль',
    'Август',
    'Сентябрь',
    'Октябрь',
    'Ноябрь',
    'Декабрь'
  ];

  const rows = [headers];

  for (let year = startYear; year <= endYear; year++) {
    for (let monthIndex = 0; monthIndex < 12; monthIndex++) {
      const monthNumber = monthIndex + 1;
      const monthName = monthNames[monthIndex];

      const quarter = 'Q' + Math.ceil(monthNumber / 3);
      const quarterName = quarter + ' ' + year;

      const monthStart = new Date(year, monthIndex, 1);
      const monthEnd = new Date(year, monthIndex + 1, 0);

      const sprints = buildMonthSprints_(monthStart, monthEnd);

      sprints.forEach((sprint, index) => {
        const sprintNumber = index + 1;
        const sprintId =
          year +
          '-' +
          String(monthNumber).padStart(2, '0') +
          '-S' +
          sprintNumber;

        const sprintName =
          monthName +
          ' / Спринт ' +
          sprintNumber +
          ' | ' +
          formatDateShort_(sprint.start) +
          '–' +
          formatDateShort_(sprint.end);

        rows.push([
          year,
          quarter,
          quarterName,
          monthNumber,
          monthName,
          sprintId,
          sprintNumber,
          sprintName,
          sprint.start,
          sprint.end,
          sprint.days
        ]);
      });
    }
  }

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#0b5d1e')
    .setFontColor('#ffffff');

  sheet.getRange(2, 9, rows.length - 1, 2).setNumberFormat('dd.mm.yyyy');
  sheet.getRange(2, 7, rows.length - 1, 1).setNumberFormat('0');

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function buildMonthSprints_(monthStart, monthEnd) {
  let sprints = [];
  let current = stripTime_(monthStart);

  while (current <= monthEnd) {
    let end = getSunday_(current);
    if (end > monthEnd) end = stripTime_(monthEnd);

    sprints.push({
      start: stripTime_(current),
      end: stripTime_(end),
      days: daysBetweenInclusive_(current, end)
    });

    current = addDays_(end, 1);
  }

  if (sprints.length > 1 && sprints[0].days < 3) {
    sprints[1].start = sprints[0].start;
    sprints[1].days = daysBetweenInclusive_(sprints[1].start, sprints[1].end);
    sprints.splice(0, 1);
  }

  if (sprints.length > 1 && sprints[sprints.length - 1].days < 3) {
    const last = sprints[sprints.length - 1];
    const prev = sprints[sprints.length - 2];

    prev.end = last.end;
    prev.days = daysBetweenInclusive_(prev.start, prev.end);

    sprints.splice(sprints.length - 1, 1);
  }

  return sprints;
}

function getSunday_(date) {
  const d = stripTime_(date);
  const day = d.getDay(); // 0 = Sunday
  const daysToSunday = day === 0 ? 0 : 7 - day;
  return addDays_(d, daysToSunday);
}

function addDays_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return stripTime_(d);
}

function stripTime_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function daysBetweenInclusive_(start, end) {
  return Math.round((stripTime_(end) - stripTime_(start)) / 86400000) + 1;
}

function formatDateShort_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM');
}

function updateSprintDatesOnEdit_(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.range.getSheet();

  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();

  const COL_SPRINT_SELECT =
    LAYOUT.SPRINT.SELECTOR.VALUE.START_COLUMN;
  const COL_DATE_VALUE =
    LAYOUT.SPRINT.DATE.VALUE.START_COLUMN;

  if (editedCol !== COL_SPRINT_SELECT) return;

  const sprintRows = getSprintInputRowsV1_();

  if (!sprintRows.includes(editedRow)) return;

  const sprintName = e.range.getValue();

  const startDateCell = sheet.getRange(editedRow, COL_DATE_VALUE);       // BF строки выбора спринта
  const endDateRow =
    editedRow +
    LAYOUT.SPRINT.ROW_OFFSETS.PROGRESS -
    LAYOUT.SPRINT.ROW_OFFSETS.SELECTOR;
  const endDateCell = sheet.getRange(endDateRow, COL_DATE_VALUE);

  if (!sprintName) {
    startDateCell.clearContent();
    endDateCell.clearContent();
    return;
  }

  const calendarSheet = ss.getSheetByName('95_CALENDAR');
  const lastRow = calendarSheet.getLastRow();

  const calendarValues = calendarSheet
    .getRange(2, 8, lastRow - 1, 3) // H:J = Sprint_Name, Sprint_Start, Sprint_End
    .getValues();

  const found = calendarValues.find(row => row[0] === sprintName);

  if (!found) {
    startDateCell.setValue('Не найдено');
    endDateCell.setValue('Не найдено');
    return;
  }

  startDateCell.setValue(found[1]);
  endDateCell.setValue(found[2]);

  startDateCell.setNumberFormat('dd.mm.yyyy');
  endDateCell.setNumberFormat('dd.mm.yyyy');
}

function getSprintInputRowsV1_() {
  const rows = [];

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
      rows.push(
        getSprintLayoutV1_(monthIndex, sprintIndex).selectorRow
      );
    }
  }

  return rows;
}
