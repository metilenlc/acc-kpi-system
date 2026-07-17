/**
   * 
   * function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu('ACC')


      .addItem('Обновить календарь спринтов', 'buildMonthlySprintCalendar_2025_2030')

      .addSeparator()

      .addItem('Пересобрать шаблон', 'buildAccTemplateV1_Layout')

      .addSeparator()

      .addItem('Создать группировку квартальных KPI и итогов', 'setupQuarterKpiSummaryGroupV1')  
      .addItem('Создать группировки месяцев', 'setupMonthGroupsV1')
      .addItem('Создать группировки KPI и итогов месяца', 'setupMonthKpiSummaryGroupsV1')
      .addItem('Создать группировки спринтов', 'setupSprintRowGroupsV1')
      .addItem('Удалить группировки строк', 'clearRowGroupsV1')


      .addSeparator()

      .addItem('Подключить списки', 'setupDropdownsV1')

      .addSeparator()

      .addItem('Удалить выпадающие списки на активном листе', 'clearDropdownsOnActiveSheet')

      .addSeparator()

      .addItem('Подключить прогресс-бары спринтов', 'setupSprintProgressBarsV1')

      .addSeparator()

      .addItem('Очистить прогресс-бары спринтов', 'clearOldSprintProgressBarsV1')

      .addSeparator()

      .addItem('Пересчитать спринты', 'updateSprintTotalsV1')
      .addItem('Пересчитать месяцы', 'updateMonthTotalsV1')
      .addItem('Пересчитать квартал', 'updateQuarterTotalsV1')
      .addItem('Пересчитать KPI месяцев', 'updateKpiMonthTotalsV1')
      .addItem('Рассчитать планы KPI месяцев', 'calculateKpiMonthPlansV1')
      .addItem('Пересчитать KPI квартала', 'updateKpiQuarterTotalsV1')
      
      .addSeparator()

      .addItem('Пересчитать всё', 'recalculateAnalyticsV1')

      .addSeparator()

      .addItem('Очистить лист', 'clearAccountDataV1_')
      

      .addToUi();

    }

  function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ACC')
    .addItem('Создать все аккаунты', 'createAccountsV1')
    .addItem('Пересчитать всё', 'recalculateAnalyticsV1')
    .addSeparator()
    .addItem('Подключить списки', 'setupDropdownsV1')
    .addItem('Создать группировки спринтов', 'setupSprintRowGroupsV1')
    .addItem('Свернуть группы', 'collapseActiveSheetGroupsV1')
    .addItem('Развернуть группы', 'expandActiveSheetGroupsV1')
    .addToUi();
  }
/*

/**
 * Пользовательское меню.
 * Всегда показывается при открытии таблицы.
 */

function onOpen() {
  buildUserMenuV1_();

  // Для режима разработки раскомментировать:
   buildDeveloperMenuV1_();
}


/**
 * Меню для пользователей.
 */
function buildUserMenuV1_() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Работа')
    .addSubMenu(
      ui.createMenu('Навигация')
        //.addItem('Реестр аккаунтов', 'openAccountsRegistryV1')
        .addItem('Реестр аккаунтов', 'openAccountsRegistryV1')
        .addItem('Справочник', 'openKpiDirectoryV1')
        .addSeparator()
        .addItem('Открыть выбранный аккаунт', 'openSelectedAccountV1')
        .addItem("Найти текущий аккаунт (в реестре)", "findCurrentAccountV1")
    )
    .addSubMenu(
      ui.createMenu('Расчёты')
        .addItem('Распределить KPI по месяцам', 'calculateKpiMonthPlansV1')
        .addItem('Пересчитать KPI', 'recalculateKpiAnalyticsV1')
        .addItem('Пересчитать гипотезы', 'recalculateSprintAnalyticsV1')
        .addItem('Полный пересчёт', 'recalculateAnalyticsV1')
        .addSeparator()
        .addItem("Пересчитать KPI всех аккаунтов", "recalculateAllAccountsKpiV1")
    )
    .addSubMenu(
      ui.createMenu('Группы')
      
        .addItem('Рабочий вид', 'showWorkViewV1')
        .addSeparator()
        .addItem('Свернуть абсолютно всё', 'collapseAllAccountGroupsV1')
        .addItem('Развернуть абсолютно всё', 'expandAllAccountGroupsV1')
        .addSeparator()
        .addItem('Показать только месяцы', 'showMonthsOnlyV1')
        .addItem('Показать месячные KPI', 'showMonthKpiOnlyV1')
        .addItem('Открыть квартальные KPI', 'expandQuarterKpiV1')
    )
    
    .addToUi();
}


/**
 * Меню разработчика.
 * Подключается вручную через buildDeveloperMenuV1_()
 * или раскомментированием вызова в onOpen().
 */
function buildDeveloperMenuV1_() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Разработка')
    .addItem('Пересобрать шаблон', 'buildAccTemplateV1_Layout')
    .addItem('Обновить календарь спринтов', 'buildMonthlySprintCalendar_2025_2030')
    .addSeparator()
    .addItem('Подключить спринты ко всем аккаунтам', 'updateSprintDropdownsForAllAccountsV1')
    .addSeparator()
    .addItem('Подключить списки', 'setupDropdownsV1')
    .addItem('Удалить списки', 'clearDropdownsOnActiveSheet')
    .addSeparator()
    .addItem('Тест: создать один лист аккаунта', 'createNextAccountSheetTestV1')
    .addItem('Пакет: создать листы аккаунтов', 'createAccountSheetsBatchV1')
    .addItem('Удалить все листы аккаунтов', 'deleteAllAccountSheetsV1')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Расчёты')
        .addItem(
          'Распределить KPI по всем аккаунтам',
          'distributeKpiPlansForAllAccountsBatchV1'
        )
    )
    .addSeparator()
    .addItem('Группировка квартала', 'setupQuarterKpiSummaryGroupV1')
    .addItem('Группировки месяцев', 'setupMonthGroupsV1')
    .addItem('Группировки KPI и итогов', 'setupMonthKpiSummaryGroupsV1')
    .addItem('Группировки спринтов', 'setupSprintRowGroupsV1')
    .addItem('Удалить группировки', 'clearRowGroupsV1')
    .addSeparator()
    .addItem('Свернуть все группы', 'collapseActiveSheetGroupsV1')
    .addItem('Развернуть все группы', 'expandActiveSheetGroupsV1')
    .addToUi();
}
