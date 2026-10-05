# Отчёт о неисправностях и ограничениях

Проект: `statistica-mcp` 2.3.0
Дата: 03.10.2026
Статус: сервер работает; все дефекты v1.0.0 исправлены, автопроверка `selftest.mjs` — 44/44. Ниже — что было исправлено, что добавлено в v2.2–v2.3 и что остаётся сознательным ограничением.

---

## 1. Исправленные дефекты (были в 1.0.0)

| ID | Серьёзность | Инструменты | Проблема | Статус | Коммит |
|---|---|---|---|---|---|
| B-1 | высокая | `descriptives`, `read_variables` | Пропуски (`-9999`) не распознавались, искажали статистики | исправлено | `bf15f92` |
| B-2 | низкая | `read_variables` | Печаталось `variable undefined` | исправлено | `bf15f92` |
| B-3 | средняя | `write_variables` | Не находил переменную по имени, добавленную в сессии | исправлено | `bf15f92` |
| B-4 | средняя | все | Служебные префиксы в именах (`!~\FT1,,,\c0,`) | исправлено (чистое имя) | `bf15f92` |
| B-5 | косметическая | `describe_spreadsheet` | `sheetIndex: 0` без явного листа | исправлено | `bf15f92` |
| B-6 | низкая | `export_csv`, `save_spreadsheet` | Файл мог быть ещё не записан | исправлено (`Wait-File`) | `bf15f92` |

### B-1. Пропуски

Теперь `Read-Variable` читает `VariableMissingData(var)` и считает пропуском **оба** маркера: `-999999998` (никогда не записанная ячейка) и объявленный код переменной (в реальных файлах встречается `-9999`). Пропуски возвращаются как `null`; `descriptives` считает их корректно.

### B-2. Индекс переменной

`Read-Variable` возвращает поле `index`; `read_variables` печатает `variable 2: SERIES_G`.

### B-3. Поиск переменной

`Get-VarIndex`: сначала `VariableNumber(name)`, при неудаче — линейный перебор по короткому и «чистому» имени. Дополнительно строковый ключ вида `"1"` трактуется как индекс (коммит `352e2a5`).

### B-4. Чистые имена

`Get-CleanName` убирает служебный префикс `!~\...` (берёт часть после последней запятой). `describe_spreadsheet` показывает и `name`, и `cleanName`.

### B-5. Индекс листа

`Open-Sheet` запоминает реальный 1-based индекс и имя; `describe_spreadsheet` возвращает корректные `sheetIndex`/`sheetName`.

### B-6. Готовность файла

`Wait-File` до 20 раз с паузой 250 мс дожидается появления файла перед чтением размера.

---

## 2. Дефекты, найденные и исправленные при доработке v2

| ID | Инструмент | Проблема | Причина | Исправление | Коммит |
|---|---|---|---|---|---|
| B-7 | `statistica_time_series` | `TypeOfTransformation` падал (`Access Violation`), сглаживание не применялось | Сеттер ждёт индекс без флага `0x40000000` | Передавать `константа − 0x40000000` | `408daea` |
| B-8 | `statistica_regression` | Регрессия не отдавала результаты через модуль 1701 | Многошаговый «мастер»; `.Summary` недоступен | Переход на модуль GRM (4601): `.Coefficients`, `.UnivariateResults` | `408daea` |
| B-9 | `statistica_regression` | Задавалось несуществующее свойство `RegressionStandard` | Неверное имя флага | Флаги только для не-`standard` методов | `408daea` |
| B-10 | `rename_variables` | Числовые ключи JSON (`"1"`) не принимались | Ключ трактовался как имя | Строка `^\d+$` → индекс | `352e2a5` |
| B-11 | `run_analysis` | `set` на каждом свойстве пересоздавал диалог | `$ana.Dialog` вызывался в цикле | Диалог кэшируется один раз на шаг | `352e2a5` |

---

## 2a. Добавлено в v2.2–v2.3 и особенности COM

| Возможность | Механизм | Замечание |
|---|---|---|
| Имена наблюдений | `Spreadsheet.CaseName(i)` (Get/Let) | Индексируемое свойство; чтение и запись — через `CallByName`. |
| Сортировка | `SortDataEx(Variables, Direction, By, false, true)` | Принимает **ровно один ключ** за вызов: массив из двух переменных даёт `Invalid variable list`. Несколько ключей — последовательными вызовами от младшего к старшему. |
| Выбор строк | `SetSize` + повторная запись всех колонок в одном сеансе COM | Маска считается в Node; воркеру передаётся готовый список 1-based номеров строк. |
| Перекодирование | чтение + отображение в Node + `write` | COM-метод `Recode` через `CallByName` падает с `E_INVALIDARG` (сложная сигнатура с массивами), поэтому реализован через данные. |
| ANOVA / GLM | модуль 4100, шаги `GLMAnalysisItem=1` → Run → `Variables="dep \| effects"` → Run | Тот же диалог, что у GRM: результаты — `UnivariateResults` (таблица ANOVA) и `Coefficients`. |
| Факторный анализ | модуль 2101: `Variables` → Run → выбрать метод (`PrincipalComponents` и др.) → Run | Двухшаговый мастер; результаты (`Eigenvalues`, `FactorLoadings`, `Communalities`) доступны после второго запуска. |
| Лаговые произведения | `addwrite` (AddVariables + запись в одном сеансе) | Имена `Lag1..LagK`; значения считаются в Node. |
| Подавление диалогов | `Application.DisplayAlert=$false` + `ExportTextEx`/`ExportXLS` + `Close($false)` | Убраны окна «features will be lost», «Save As Text File», «Save changes to Workbook1?», блокировавшие `Quit`. |
| Уровни измерения | `VariableMeasurementType(idx, type)` | `continuous/categorical/ordinal`; используется как фактор/ковариата в GLM. |
| Метки значений | `SetTextLabel`, `TextLabelValue`, `NumberOfTextLabels` | Чтение/запись; `RemoveAllLabels` сбрасывает. |
| PDF-экспорт | PNG → PDF своим конвертером (zlib) | `SaveAsPDF`/`SaveAsFormat(PDF)` отдают `False`; PDF с валидным `xref`. |
| Кластерный анализ | модуль 2201, `Run` → `Variables` → `Run` | `AmalgamationSchedule`, `DistanceMatrix`, `DescriptiveStatistics`. |
| Список листов | `Application.Spreadsheets` | `list_sheets`; открытие `.stw` с несколькими листами по имени/индексу. |
| Скриншот окна | `CopyFromScreen` + `Activate()` документа | `PrintWindow` не рисует дочерние MDI-окна; `mode=document` снимает таблицу/график. |
| Повтор COM | `New-ComApp` (3 попытки) | Зависший экземпляр даёт разовый `E_FAIL`; создание повторяется. |
| DPI-awareness | `SetProcessDPIAware` в начале воркера | Без него `Screen.Bounds`/`CopyFromScreen` давали 2560×1359 вместо 5120×2718 — кадр обрезался. |
| Активация окна | `ShowWindow` (restore+maximize), `AppActivate`, `SwitchToThisWindow` | `SetForegroundWindow` игнорируется для свёрнутого окна; снимок — весь `VirtualScreen`, ответ предупреждает, если фокус не получен. |
| Запаздывание ряда | `ShiftSeriesForward` + `LagForShiftingSeriesForward` | `statistica_time_series procedure=shift` (есть и `direction=back`). |
| Центрированное SMA | `NPointsMovingAverage` + `ComputeMovingAverageFromPriorValues=false` | Окно = период; `prior=true` — нецентрированное. |
| Fit-линия | МНК в Node → `addwrite` | `add_fit_line`: fitted-значения пишутся переменной, рисуются рядом с данными. |
| Совмещённый и 3D-график | `properties.GraphType` | 11012 `GraphType=1` — один график с несколькими линиями; 11021 `GraphType=6` — Surface. |
| TS-мультиграфик | `PlotMultipleVariables` | *Отвергнуто:* сеттер роняет STATISTICA (RPC unavailable); обход — 11012 `GraphType=1`. |
| SVB-макрос | `Application.Open(.svb)` + `Macro.Execute()` | `run_macro`; перед запуском `$ss.Activate()` (иначе `ActiveSpreadsheet` = Nothing); макросы закрываются в `Close-AllDocuments`. |
| Экспон. сглаживание | `NoTrend*`/`LinearTrend*`/`Damped*` + `ParameterAlpha/Gamma/Delta` | `exponential_smoothing` с `model` (ЛР4: EMA/Хольт/Тейл–Вейдж/Уинтерс). |
| Полиномиальный fit | МНК в Node (нормальные уравнения, метод Гаусса) | `add_fit_line degree=1..6`. |
| Нормальность | `ShapiroWilkWTest`, `KSAndLillieforsTestForNormality` | `statistica_normality` (1301); `Histograms` — метод, не флаг. |
| Пошаговая матрица | чтение ряда + `addwrite` колонки | `add_lag_column` дописывает одну колонку `Lag_m` (product/shift + центрир. SMA). |
| Снимок панели модуля | скрытое окно `#32770` + `ShowWindow`/`CopyFromScreen` | `statistica_dialog`; окно на передний план, режим `screen`/`dialog`; закрывается только диалог. |
| Импорт в окно | `ImportTextAutoEx` (attach) | лист импорта эфемерный: нужно `save` и открыть файлом, иначе `Spreadsheets` его не видит. |
| Сохранение в окне | `save_as` + `attach` | `save_spreadsheet` с `copy=false` сохраняет открытый лист на месте. |
| Постоянное окно | `GetActiveObject`/`New-Object`, `Visible=$true`, без `Quit` | `statistica_open`: окно не закрывается; при нескольких экземплярах `GetActiveObject` отдаёт первый, поэтому `open` переиспользует запущенный. |
| `attach` без `path` | пустой `path` → `ActiveSpreadsheet` | `requirePath` разрешает пустой путь в `attach`. |

### Почему правка данных делается в одном вызове воркера

Режим stateless: каждый вызов поднимает и закрывает новый `statist.exe`, состояние между вызовами не сохраняется. Поэтому `select` (уменьшение числа строк и перезапись колонок) и `sort` (устойчивое изменение порядка) выполняются **одной** командой воркера, а не композицией серверных шагов. `recode` — исключение: он читает и записывает одну и ту же переменную разными процессами, что корректно, т.к. запись полностью замещает колонку.

---

## 3. Проверено и работает

| Операция | Вызов | Результат |
|---|---|---|
| Открытие `.sta`/`.stw` | `Application.Open` + `Spreadsheets.Item` | лист и переменные корректны |
| Чтение колонки | `get_VData(n)` / `GetData(..., LabelsAsText)` | числа/строки, пропуски → `null` |
| Запись колонки | `put_VData(n, [double[]])` | векторно |
| Учёт пропусков | `get_VariableMissingData(n)` | `-9999` и `-999999998` |
| Создание/удаление/переименование | `AddVariables`, `DeleteVariables`, `set_VariableName` | работает |
| Описательные статистики | модуль 1301, `Statistics = scBasDescriptives` | таблица статистик |
| Корреляция | модуль 1301, `Statistics = scBasCorrelationMatrices` | матрица |
| t-тест | модуль 1301, `Statistics = 5/6` | `.Summary`, `.TTests` |
| Регрессия | модуль 4601 (GRM) | `.Coefficients`, `.UnivariateResults` |
| ACF/PACF | модуль 1901 | `.Autocorrelations`, `.PartialAutocorrelations` |
| ARIMA | модуль 1901 | `.Summary`, `.ForecastCases` |
| Спектр (Фурье) | модуль 1901, `SpectralFourierAnalysis` | `.Summary` (Frequency/Period/Coefficients) |
| Сглаживание | модуль 1901, `TypeOfTransformation = 1` | `.SaveVariables` (исходный + сглаженный) |
| Импорт текста/Excel | `ImportTextAuto*`, `ImportXLSAsSpreadsheet` | новые листы |
| Экспорт/сохранение | `SaveAs`, `SaveCopyAs` | `.csv`, `.sta`, `.xlsx` |
| Запись формул | `VariableLongName(idx)` + `Recalculate(idx)` | `=v2*2` даёт 2× значений, longName хранит формулу |
| Экспорт графиков | `Graph.SaveAs(path)` | изображение (PNG с корректной сигнатурой) или `.stg` |
| Имена наблюдений | `CaseName(i)` Get/Let | чтение и запись |
| Сортировка | `SortDataEx` по одному ключу | порядок меняется, имена строк сохраняются |
| Выбор строк | `SetSize` + запись | 144 из 156 (по `notmissing`) |
| Перекодирование | read → map → write | 2 значения изменены в тесте |
| ANOVA / GLM | модуль 4100 | `UnivariateResults`, `Coefficients` |
| Факторный анализ | модуль 2101 | `Eigenvalues`, `FactorLoadings`, `Communalities` |
| Лаговые произведения | AddVariables + запись | `Lag1..LagK` |

---

## 4. Сознательно не реализовано

### Шаблоны графиков `.stg`
Сохранение графа в `.stg` **поддерживается**: `saveGraph` (и параметр `out` у `statistica_time_series`) пишет файл через `Graph.SaveAs` — расширение задаёт формат (`.png/.jpg/.emf` — изображение, `.stg` — родной граф STATISTICA); из результата-массива сохраняются только графики. А вот применения `.stg` как шаблона стиля к другому графу в COM нет: `.stg` открывается как граф-документ (`Application.Open`), но `Plot.Style` не присваивается («не удалось задать свойство Style для каждого объекта»), `Plot.CopyToStyle(Object)` требует недоступного аргумента. Вместо шаблона используйте необязательный `properties` у `statistica_graph` — он задаёт параметры диалога графа.

### Экспорт PDF (растянутый на страницу, не векторный)
Своего PDF STATISTICA в headless не пишет (`SaveAsPDF`/`SaveAsFormat(PDF)` → `False`), поэтому PDF собирается в Node из PNG-снимка графа (растровый одностраничный PDF). Векторный PDF через драйвер печати не поддержан.

### Долгоживущая сессия
Каждый вызов = новый процесс `statist.exe` (stateless); изменения на диске — только с явным `save`. Режим `attach` работает с уже открытым окном без закрытия. Демон для снижения задержки — P2.

### Методы с `out`-параметрами
`Statistics`, `ColumnStats`, `Recode`, `FilterCases` через `CallByName` недоступны (сложные сигнатуры/`DISP_E_*`). Соответствующие операции реализованы через данные: статистики — в Node, `recode` — read/map/write, `select_cases` — маска в Node + перезапись колонок.

### Полнота пресетов
Пресеты покрывают описательные, корреляцию, частоты, t-тест, регрессию, ANOVA/GLM, факторный и кластерный анализ, временные ряды. Остальные модули (дискриминантный, нейросети, SEPATH и т.д.) доступны через `run_analysis` + `describe_analysis`.

---

## 5. План развития

- Долгоживущий `statist.exe` (демон) — снижение задержки на вызов (P2).
- Векторный PDF (через драйвер печати) вместо растрового (P2).
- Применение шаблонов графиков `.stg`, если найдётся рабочий COM-путь (P2).
- Дискриминантный анализ отдельным пресетом (по запросу).
