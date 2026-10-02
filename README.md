# STATISTICA MCP

MCP-сервер для полноценного управления STATISTICA 12 через её COM-интерфейс. Позволяет агенту (opencode, Claude Desktop и т.п.) работать с файлами `.sta`/`.stw`, редактировать данные, **запускать собственные процедуры STATISTICA** (базовая статистика, корреляция, t-тесты, регрессия, временные ряды и любые другие модули) и забирать результаты таблицами.

**Ноль внешних зависимостей.** Node.js реализует протокол MCP сам, PowerShell-worker обращается к COM. `npm install` не нужен.

---

## Требования

| Компонент | Версия | Проверено |
|---|---|---|
| Windows | 10/11 | да |
| Node.js | >= 18 | 24.19.0 |
| STATISTICA 12 | любая | 12.5.192.7 (x64) |
| PowerShell | 5.1 (системный) | да |

Интернет не требуется.

Проверить, что COM доступен:

```powershell
$app = New-Object -ComObject "STATISTICA.Application"
$app.Visible = $false
$app.Version
$app.Quit()
```

---

## Установка

1. Распаковать проект в любую папку (например `C:\tools\statistica-mcp`).
2. Добавить сервер в `~/.config/opencode/opencode.jsonc`:

```jsonc
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "statistica": {
      "type": "local",
      "command": ["node", "C:\\tools\\statistica-mcp\\statistica\\server.mjs"],
      "enabled": true
    }
  }
}
```

Для Claude Desktop — `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "statistica": {
      "command": "node",
      "args": ["C:\\tools\\statistica-mcp\\statistica\\server.mjs"]
    }
  }
}
```

3. Перезапустить клиент.

> Путь в конфиге — именно на `server.mjs`. Он определяет расположение `sta.ps1` относительно себя, поэтому папку можно переносить целиком, но нельзя отделять `sta.ps1` от `server.mjs`.

---

## Проверка установки

В папке проекта:

```bash
node selftest.mjs "C:\путь\к\LAB3.sta"
```

Скрипт создаёт временную копию файла, прогоняет **27 проверок** (все инструменты, включая анализы и импорт) и печатает результат. Исходный файл не изменяется.

---

## Архитектура

```
агент  ──JSON-RPC 2.0 (stdio)──▶  server.mjs  ──запуск процесса──▶  sta.ps1  ──COM──▶  statist.exe
```

- **`server.mjs`** — MCP-сервер. Разбор протокола, схемы инструментов, таблицы модулей и enum-констант, форматирование результатов. На stdout пишет только JSON-RPC, все логи в stderr.
- **`sta.ps1`** — воркер. Читает JSON-запрос из файла, делает вызовы COM, пишет JSON-ответ в файл. Обмен через файлы снимает проблемы с кодировками кириллицы в PowerShell 5.1.
- Каждый вызов инструмента порождает **новый** процесс `statist.exe` и закрывает его. Состояние между вызовами не сохраняется (stateless), поэтому изменения на диске требуют явного параметра `save`.
- Один вызов = один файл `.sta`.

---

## Инструменты

### Общие

| Инструмент | Назначение |
|---|---|
| `statistica_info` | Доступность COM, версия, путь к `statist.exe`, PID. Вызывайте первым при сбоях. |
| `list_analysis_modules` | Список всех модулей анализа (id + имя) для `run_analysis`. |
| `describe_spreadsheet` | Открыть `.sta`/`.stw`: размер, список переменных (индекс, короткое/чистое/длинное имя, тип, уровень измерения, код пропуска). |
| `read_variables` | Чтение данных. Числовые колонки — векторно, текстовые — строками; пропуски → `null`. |
| `write_variables` | Полная перезапись переменных; числовая колонка требует ровно по значению на наблюдение. |
| `add_variables` | Добавление пустых переменных (`0` numeric, `1` text, `2` integer, `3` byte). |
| `rename_variables` | Переименование коротких/длинных имён (по имени или индексу). |
| `delete_variables` | Удаление диапазона переменных. |
| `set_size` | Изменение размера таблицы. |
| `import_data` | Импорт текста/Excel в STATISTICA. |
| `export_csv` | Экспорт листа штатным CSV-писателем. |
| `save_spreadsheet` | Сохранение листа в новый файл (`.sta`, `.stw`, `.csv`, `.xlsx`). |

### Статистика

| Инструмент | Модуль / механизм |
|---|---|
| `descriptives` | Быстрый расчёт в Node по данным из COM (N, missing, mean, sd, se, min, q1, median, q3, max, sum). |
| `statistica_descriptives` | Описательные статистики **движком** Basic Statistics (включая квантили, асимметрию, эксцесс). |
| `statistica_correlation` | Матрица корреляций Пирсона. |
| `statistica_frequencies` | Частотные таблицы и гистограммы. |
| `statistica_t_test` | t-тесты: `single` (к константе) и `dependent` (парные). |
| `statistica_regression` | Множественная регрессия (модуль GRM). |
| `statistica_time_series` | Временные ряды: `descriptives`, `autocorrelation`, `partial_autocorrelation`, `cross_correlation`, `arima`, `spectral`, `smoothing`, `exponential_smoothing`, `differencing`, `seasonal_decomposition`. |

### Универсальный движок

| Инструмент | Назначение |
|---|---|
| `describe_analysis` | Интроспекция диалога анализа: свойства, методы и известные enum-константы. Помогает собрать запрос к `run_analysis`. |
| `run_analysis` | Запуск **любой** процедуры последовательностью шагов. |

#### Шаги `run_analysis`

```jsonc
{
  "path": "C:\\data\\LAB3.sta",
  "module": 1901,
  "steps": [
    { "set":  { "Variables": "2", "FocusTimeSeriesVariable": 1 } },
    { "call": "ARIMAAndAutocorrelationFunctions" },
    { "set":  { "NumberOfAutoregressiveParameters": 1, "NumberOfMovingAverageParameters": 0 } },
    { "run":  true },
    { "set":  { "NumberOfCasesToForecast": 12 } },
    { "result": "Summary" },
    { "result": "ForecastCases" }
  ]
}
```

- `set` — присвоить свойства диалога (массивы и строки допустимы; `Variables` принимает `"1 3 5"`, `"2-4"` или массив индексов).
- `call` — вызвать метод диалога (например, `SpectralFourierAnalysis`, `Transformations`, `ExponentialSmoothingAndForecasting`). Опционально `args`.
- `run` — выполнить анализ (`Application.Analysis(...).Run`).
- `result` — прочитать свойство после запуска. Результат маршалится как таблица, массив, коллекция или идентификатор документа. Если свойство параметризованное (например `Summary(1)`), индекс подставляется автоматически.

Неверные свойства/методы не прерывают анализ: они попадают в блок `Warnings` ответа.

---

## Примеры

**Описательные и корреляция**

```
statistica_descriptives { "path": "...\\LAB3.sta", "variables": [2, 3, 4] }
statistica_correlation   { "path": "...\\LAB3.sta", "variables": [2, 3, 5] }
```

**Параметрическая регрессия**

```
statistica_regression { "path": "...\\LAB3.sta", "dependent": 2, "predictors": [1] }
```

**ARIMA с прогнозом**

```
statistica_time_series {
  "path": "...\\LAB3.sta", "procedure": "arima", "variables": [2],
  "arOrder": 1, "maOrder": 0, "difference": true, "forecasts": 12
}
```

**Сглаживание и спектр**

```
statistica_time_series { "path": "...", "procedure": "smoothing", "variables": [2], "window": 3 }
statistica_time_series { "path": "...", "procedure": "spectral",  "variables": [2] }
```

**Любая процедура — движок**

```
list_analysis_modules {}
describe_analysis { "path": "...\\LAB3.sta", "module": 4100 }   // GLM
run_analysis {
  "path": "...\\LAB3.sta", "module": 1301,
  "steps": [
    { "set": { "Statistics": 0 } },
    { "run": true },
    { "set": { "Variables": "2-4", "Mean": true, "ValidN": true } },
    { "result": "Summary" }
  ]
}
```

---

## Особенности COM, учтённые сервером

- **ProgID `STATISTICA.Application` работает, а `_Application` — нет.** Сборка interop не используется, всё через `New-Object -ComObject` + `Microsoft.VisualBasic.Interaction::CallByName`.
- **Параметризованные свойства** (`Data(case,var)`, `VData(var)`, `VariableName(n)`) требуют `CallByName` с `CallType::Get` / `::Let`.
- **Массивы маршалятся только типизированные.** `put_VData(n, [double[]])` записывает колонку векторно; длина массива должна быть ровно `NumberOfCases`.
- **Пропуски.** `VData` возвращает `-999999998` для «никогда не записанных» ячеек, а у переменной может быть собственный код (`VariableMissingData`, в реальных файлах встречается `-9999`). Сервер считает пропуском оба.
- **`GetData` при `nbVars > 1` добавляет метку строки**, поэтому колонки читаются по одной.
- **`CallByName` не сочетается с массивом значений в одном аргументе** — из-за этого используется `comSetOne` и ручная сборка `[object[]]`.
- **Time Series `TypeOfTransformation`.** Сеттер не принимает флаговые константы (`0x40000000 + n`) и падает с `Access Violation`; нужно передавать индекс без флага: `Smoothing = 1`, `Fourier = 5`, `Autocorrelation = 7`, `Descriptive = 8`, `Differencing = 4`. Таблица — в `describe_analysis`.
- **ARIMA `NumberOfCasesToForecast`** задаётся после `Run`.
- **`SaveAsPDF` зависает** в headless-режиме — не используется.
- **Методы с `out`-параметрами** (`Statistics`, `ColumnStats`) через `CallByName` не работают, поэтому часть описательных статистик считает Node.

---

## Ограничения

1. **Stateless:** каждый вызов — новый процесс `statist.exe`, задержка в несколько секунд; изменения сохраняются только с `save`.
2. **Графики** возвращаются как идентификаторы документов, экспорт изображений не реализован.
3. **Запись формул** STATISTICA не удалась; производный ряд считается в Node и пишется значениями.
4. **Пресеты** покрывают популярные сценарии; полный доступ — через `run_analysis` + `describe_analysis`.

---

## Доработка

- **Новый инструмент:** добавьте объект в массив `TOOLS` в `server.mjs` и ветку в `switch (cmd)` файла `sta.ps1`.
- **Новый анализ:** обычно достаточно `run_analysis`; пресет нужен только для частого сценария.
- **Известные enum-константы** перечислены в `ENUMS` (`server.mjs`); их можно дополнить из отчёта.

---

## Файлы

```
statistica/
  server.mjs                     MCP-сервер, протокол, инструменты
  sta.ps1                        COM-воркер (файловый обмен)
  selftest.mjs                   самопроверка (27 проверок)
  package.json                   зависимостей нет
  reports/
    development-report.md        отчёт: проблемы, гипотезы, решения, итог
```
