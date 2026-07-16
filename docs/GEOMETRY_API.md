# Geometry API

Инвентаризация функций, объявленных в `03_Geometry.js`. Под «Public» понимается функция без завершающего `_`, доступная как точка входа Apps Script; функции с завершающим `_` классифицированы как внутренние. «Can be deprecated» указывает, можно ли вывести текущую функцию из употребления после миграции её потребителей или переноса ответственности.

## `getActiveAccountSheetV1_`

- **Параметры:** нет.
- **Возвращаемое значение:** активный объект `GoogleAppsScript.Spreadsheet.Sheet`; выбрасывает ошибку, если открыт `ACCOUNTS` или лист, имя которого не начинается с `ACC_`.
- **Назначение:** получает и проверяет активный рабочий лист аккаунта; `ACC_TEMPLATE` также считается допустимым.
- **Called from:** `05_Dropdowns.js`, `07_Calculations.js`, `13_Groups.js`.
- **Public / Internal:** Internal.
- **Can be deprecated:** No — функция активно используется; при этом её следует перенести из Geometry в модуль доступа/валидации листов без изменения контракта.

## `openKpiDirectoryV1`

- **Параметры:** нет.
- **Возвращаемое значение:** `undefined`; активирует лист KPI и выделяет `A1`, а при отсутствии листа показывает предупреждение и завершает работу.
- **Назначение:** UI-команда открытия справочника KPI.
- **Called from:** `09_Menu.js` (имя функции передаётся в пункт меню «Справочник»).
- **Public / Internal:** Public.
- **Can be deprecated:** No — это активная точка входа меню; ответственность следует перенести из Geometry в UI/Directory-модуль с сохранением публичного обработчика или совместимого адаптера.

## `getKpiDictionaryV1_`

- **Параметры:** нет.
- **Возвращаемое значение:** объект-словарь, где ключ — название KPI, а значение имеет поля `unit`, `aggregation`, `distribution`, `k1`, `k2`, `k3`; при пустом справочнике возвращает `{}`.
- **Назначение:** лениво читает справочник KPI с листа `SHEETS.DIRECTORY`, нормализует значения и кэширует результат в `KPI_DICTIONARY_V1`.
- **Called from:** `07_Calculations.js`.
- **Public / Internal:** Internal.
- **Can be deprecated:** No — функция имеет действующего потребителя; её следует перенести из Geometry в Directory/Data Access-модуль.

## `getSprintStepV1_`

- **Параметры:** нет.
- **Возвращаемое значение:** число строк в одном блоке спринта: `LAYOUT.SPRINT.HEADER_ROWS + LAYOUT.SPRINT.HYPOTHESIS_COUNT + LAYOUT.SPRINT.SUMMARY_ROWS`.
- **Назначение:** вычисляет вертикальный шаг между соседними спринтами.
- **Called from:** `03_Geometry.js`: `getSprintFirstHypRowV1_`, `getSprintLayoutV1_`.
- **Public / Internal:** Internal.
- **Can be deprecated:** No — это базовый внутренний примитив канонической геометрии спринта.

## `getMonthFirstRowV1_(monthIndex)`

- **Параметры:** `monthIndex` — нулевой индекс месяца.
- **Возвращаемое значение:** номер первой строки блока месяца.
- **Назначение:** вычисляет верхнюю строку месяца по `LAYOUT.MONTH.FIRST_ROW` и `LAYOUT.MONTH.HEIGHT`.
- **Called from:** `03_Geometry.js`: `getMonthFirstKpiRowV1_`, `getMonthFirstSprintRowV1_`.
- **Public / Internal:** Internal.
- **Can be deprecated:** Yes — отдельный helper дублирует вычисление `monthRow` в `getMonthLayoutV1_` и не имеет внешних потребителей.

## `getMonthFirstKpiRowV1_(monthIndex)`

- **Параметры:** `monthIndex` — нулевой индекс месяца.
- **Возвращаемое значение:** номер первой строки KPI выбранного месяца.
- **Назначение:** прибавляет `LAYOUT.MONTH.KPI_OFFSET` к первой строке месяца.
- **Called from:** внешних и внутренних вызовов нет.
- **Public / Internal:** Internal.
- **Can be deprecated:** Yes — значение уже доступно как `kpiRow` результата `getMonthLayoutV1_`.

## `getMonthFirstSprintRowV1_(monthIndex)`

- **Параметры:** `monthIndex` — нулевой индекс месяца.
- **Возвращаемое значение:** номер первой строки первого спринта выбранного месяца.
- **Назначение:** прибавляет `LAYOUT.MONTH.SPRINT_OFFSET` к первой строке месяца.
- **Called from:** `03_Geometry.js`: `getSprintFirstHypRowV1_`.
- **Public / Internal:** Internal.
- **Can be deprecated:** Yes — значение уже доступно как `firstSprintRow` результата `getMonthLayoutV1_`.

## `getSprintFirstHypRowV1_(monthIndex, sprintIndex)`

- **Параметры:** `monthIndex` — нулевой индекс месяца; `sprintIndex` — нулевой индекс спринта.
- **Возвращаемое значение:** номер первой строки гипотез выбранного спринта.
- **Назначение:** вычисляет начало таблицы гипотез внутри блока спринта.
- **Called from:** `03_Geometry.js`: `getSprintTotalsRowV1_`, `getSprintProgressRowV1_`.
- **Public / Internal:** Internal.
- **Can be deprecated:** Yes — значение уже доступно как `firstHypRow` результата `getSprintLayoutV1_`.

## `getSprintTotalsRowV1_(monthIndex, sprintIndex)`

- **Параметры:** `monthIndex` — нулевой индекс месяца; `sprintIndex` — нулевой индекс спринта.
- **Возвращаемое значение:** номер строки итоговых значений выбранного спринта.
- **Назначение:** вычисляет строку итогов после строк гипотез и служебных строк сводки.
- **Called from:** внешних и внутренних вызовов нет.
- **Public / Internal:** Internal.
- **Can be deprecated:** Yes — значение уже доступно как `totalsRow` результата `getSprintLayoutV1_`.

## `getSprintProgressRowV1_(monthIndex, sprintIndex)`

- **Параметры:** `monthIndex` — нулевой индекс месяца; `sprintIndex` — нулевой индекс спринта.
- **Возвращаемое значение:** номер строки прогресс-бара выбранного спринта.
- **Назначение:** вычисляет строку прогресса относительно первой строки гипотез.
- **Called from:** внешних и внутренних вызовов нет.
- **Public / Internal:** Internal.
- **Can be deprecated:** Yes — значение уже доступно как `progressRow` результата `getSprintLayoutV1_`.

## `getMonthLayoutV1_(monthIndex)`

- **Параметры:** `monthIndex` — нулевой индекс месяца (`0..2` для текущего layout).
- **Возвращаемое значение:** объект `{ monthRow, titleRow, kpiRow, summaryRow, firstSprintRow }` с номерами ключевых строк месяца.
- **Назначение:** предоставляет составную геометрию выбранного месяца через единый вызов.
- **Called from:** `03_Geometry.js`: `getSprintLayoutV1_`; `04_Build.js`; `05_Dropdowns.js`.
- **Public / Internal:** Internal.
- **Can be deprecated:** No — функция является кандидатом на канонический API месяца.

## `getSprintLayoutV1_(monthIndex, sprintIndex)`

- **Параметры:** `monthIndex` — нулевой индекс месяца (`0..2`); `sprintIndex` — нулевой индекс спринта (`0..4`).
- **Возвращаемое значение:** объект `{ sprintTop, selectorRow, progressRow, headerRow, firstHypRow, lastHypRow, summaryGapRow, totalsHeaderRow, totalsRow, bottomGapRow }` с номерами ключевых строк спринта.
- **Назначение:** предоставляет составную геометрию выбранного спринта через единый вызов.
- **Called from:** `04_Build.js`, `05_Dropdowns.js`, `07_Calculations.js`.
- **Public / Internal:** Internal.
- **Can be deprecated:** No — функция является кандидатом на канонический API спринта.

## Dependency graph

Ниже указаны модули-потребители каждой функции; зависимости внутри `03_Geometry.js` показаны явно.

```text
getActiveAccountSheetV1_
├── 05_Dropdowns.js
├── 07_Calculations.js
└── 13_Groups.js

openKpiDirectoryV1
└── 09_Menu.js

getKpiDictionaryV1_
└── 07_Calculations.js

getSprintStepV1_
└── 03_Geometry.js
    ├── getSprintFirstHypRowV1_
    └── getSprintLayoutV1_

getMonthFirstRowV1_
└── 03_Geometry.js
    ├── getMonthFirstKpiRowV1_
    └── getMonthFirstSprintRowV1_

getMonthFirstKpiRowV1_
└── нет потребителей

getMonthFirstSprintRowV1_
└── 03_Geometry.js
    └── getSprintFirstHypRowV1_

getSprintFirstHypRowV1_
└── 03_Geometry.js
    ├── getSprintTotalsRowV1_
    └── getSprintProgressRowV1_

getSprintTotalsRowV1_
└── нет потребителей

getSprintProgressRowV1_
└── нет потребителей

getMonthLayoutV1_
├── 03_Geometry.js
│   └── getSprintLayoutV1_
├── 04_Build.js
└── 05_Dropdowns.js

getSprintLayoutV1_
├── 04_Build.js
├── 05_Dropdowns.js
└── 07_Calculations.js
```

## Migration candidates

Канонический Geometry API целесообразно строить вокруг составных функций, возвращающих полное описание блока:

1. **`getMonthLayoutV1_(monthIndex)`** — каноническая точка получения геометрии месяца. Она уже используется внешними модулями и должна заменить узкие функции `getMonthFirstRowV1_`, `getMonthFirstKpiRowV1_` и `getMonthFirstSprintRowV1_`.
2. **`getSprintLayoutV1_(monthIndex, sprintIndex)`** — каноническая точка получения геометрии спринта. Она уже используется внешними модулями и должна заменить `getSprintFirstHypRowV1_`, `getSprintTotalsRowV1_` и `getSprintProgressRowV1_`.
3. **`getSprintStepV1_()`** — оставить внутренним примитивом Geometry, а не самостоятельной точкой API: он обеспечивает единый расчёт высоты/шага спринта из `LAYOUT`.

Перед закреплением API рекомендуется заменить арифметические смещения внутри `getSprintLayoutV1_` (`+ 2`, `+ 4`, `+ 6` и производные) на соответствующие ключи `LAYOUT.SPRINT.ROW_OFFSETS` и декларативные размеры. Это необходимо для соблюдения правила «Geometry only through LAYOUT» и устранения magic numbers.

Функции `getActiveAccountSheetV1_`, `openKpiDirectoryV1` и `getKpiDictionaryV1_` не являются Geometry API. Их следует мигрировать соответственно в модуль работы с аккаунт-листами, UI/Directory и слой доступа к справочнику; текущие контракты нужно сохранить до перевода всех потребителей.

---

## Границы ответственности Geometry

### Принятое решение

Модуль `03_Geometry.js` должен содержать только функции, которые преобразуют
декларативную структуру `LAYOUT` в реальные координаты строк, столбцов и блоков.

Модуль Geometry не должен:

- определять активный рабочий лист;
- проверять тип активного листа;
- читать справочники;
- кэшировать данные справочников;
- переключать активные листы;
- показывать сообщения пользователю;
- выполнять бизнес-расчёты;
- изменять структуру Google Sheets.

---

## Распределение функций по модулям

| Текущая функция                                    | Целевой модуль     | Решение                | Причина                                                 |
| -------------------------------------------------- | ------------------ | ---------------------- | ------------------------------------------------------- |
| `getSprintStepV1_()`                               | `03_Geometry.js`   | Оставить               | Вычисляет вертикальный шаг спринта из `LAYOUT`          |
| `getMonthLayoutV1_(monthIndex)`                    | `03_Geometry.js`   | Оставить               | Возвращает составную геометрию месяца                   |
| `getSprintLayoutV1_(monthIndex, sprintIndex)`      | `03_Geometry.js`   | Оставить               | Возвращает составную геометрию спринта                  |
| `getMonthFirstRowV1_(monthIndex)`                  | `03_Geometry.js`   | Удалить после миграции | Дублирует поле `monthRow` из `getMonthLayoutV1_()`      |
| `getMonthFirstKpiRowV1_(monthIndex)`               | `03_Geometry.js`   | Удалить после миграции | Не имеет потребителей и дублирует `kpiRow`              |
| `getMonthFirstSprintRowV1_(monthIndex)`            | `03_Geometry.js`   | Удалить после миграции | Дублирует `firstSprintRow`                              |
| `getSprintFirstHypRowV1_(monthIndex, sprintIndex)` | `03_Geometry.js`   | Удалить после миграции | Дублирует `firstHypRow`                                 |
| `getSprintTotalsRowV1_(monthIndex, sprintIndex)`   | `03_Geometry.js`   | Удалить после миграции | Не имеет потребителей и дублирует `totalsRow`           |
| `getSprintProgressRowV1_(monthIndex, sprintIndex)` | `03_Geometry.js`   | Удалить после миграции | Не имеет потребителей и дублирует `progressRow`         |
| `getActiveAccountSheetV1_()`                       | `14_Validation.js` | Перенести              | Проверяет контекст активного листа и выбрасывает ошибку |
| `getKpiDictionaryV1_()`                            | `15_Directory.js`  | Перенести              | Читает и кэширует данные справочника KPI                |
| `openKpiDirectoryV1()`                             | `09_Menu.js`       | Перенести              | Является пользовательской командой навигации            |

---

## Канонический Geometry API

После завершения миграции в `03_Geometry.js` должны остаться следующие
канонические функции:

```javascript
getSprintStepV1_();

getMonthLayoutV1_(monthIndex);

getSprintLayoutV1_(monthIndex, sprintIndex);
```

Допускается добавление новых функций только в том случае, если они:

1. получают декларативные значения из `LAYOUT`;
2. вычисляют координаты или размеры;
3. не читают содержимое Google Sheets;
4. не записывают данные;
5. не показывают UI;
6. не выполняют бизнес-расчёты.

---

## Целевые модули

### `03_Geometry.js`

Ответственность:

- вычисление координат;
- вычисление размеров;
- построение геометрии месяца;
- построение геометрии спринта;
- проверка допустимости индексов месяца и спринта.

Не содержит обращений к:

```javascript
SpreadsheetApp.getActive();
SpreadsheetApp.getUi();
sheet.getDataRange();
sheet.getValues();
```

---

### `14_Validation.js`

Ответственность:

- проверка активного листа;
- проверка, что лист является аккаунтом;
- проверка допустимых индексов;
- проверка обязательных данных.

После переноса здесь должна находиться функция:

```javascript
getActiveAccountSheetV1_();
```

Имя функции временно сохраняется для совместимости с существующими
потребителями.

---

### `15_Directory.js`

Ответственность:

- чтение справочников;
- нормализация данных справочников;
- кэширование справочников;
- сброс кэша.

После переноса здесь должна находиться функция:

```javascript
getKpiDictionaryV1_();
```

Имя функции временно сохраняется для совместимости.

---

### `09_Menu.js`

Ответственность:

- создание меню;
- обработка пользовательских команд меню;
- открытие служебных листов.

После переноса здесь должна находиться функция:

```javascript
openKpiDirectoryV1();
```

Публичное имя функции сохраняется, поскольку оно используется Google Apps
Script как строковый обработчик пункта меню.

---

## Порядок миграции

Перенос выполняется в следующем порядке:

1. Перенести `getActiveAccountSheetV1_()` в `14_Validation.js`.
2. Проверить Dropdowns, Calculations и Groups.
3. Перенести `getKpiDictionaryV1_()` в `15_Directory.js`.
4. Проверить расчёт распределения KPI.
5. Перенести `openKpiDirectoryV1()` в `09_Menu.js`.
6. Проверить пункт меню «Справочник».
7. Перевести Geometry на `LAYOUT.SPRINT.ROW_OFFSETS`.
8. Удалить неиспользуемые узкие функции Geometry.
9. Выполнить финальный поиск ссылок по всему проекту.

---

## Ограничения миграции

На одном этапе переносится только одна функция или одна группа тесно связанных
функций.

Запрещается одновременно:

- переносить функции;
- переименовывать публичный API;
- менять бизнес-логику;
- менять структуру листа;
- удалять совместимость.

Сначала переносится функция без изменения её контракта. Переименование или
упрощение API выполняется отдельной задачей после проверки всех потребителей.

---

## Definition of Done

Разделение Geometry считается завершённым, если:

- `03_Geometry.js` содержит только вычисление геометрии;
- `getActiveAccountSheetV1_()` находится в `14_Validation.js`;
- `getKpiDictionaryV1_()` находится в `15_Directory.js`;
- `openKpiDirectoryV1()` находится в `09_Menu.js`;
- все существующие вызовы продолжают работать;
- лишние узкие функции удалены после проверки ссылок;
- в Geometry отсутствуют UI, чтение справочников и проверка активного листа.
