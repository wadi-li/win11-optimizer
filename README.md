# Win11 Optimizer — загрузка, ОЗУ и батарея (Windows 11 Pro 26H2)

Автономный HTML-анализатор и PowerShell-программа для оценки влияния автозагрузки, служб и фоновых процессов на время загрузки, расход оперативной памяти и время работы от батареи.

- **Онлайн-анализатор:** https://wadi-li.github.io/win11-optimizer/
- **Скрипт:** [`Win11-BootRamBattery-Toolkit.ps1`](Win11-BootRamBattery-Toolkit.ps1) (v2.0.0, PowerShell 5.1+)

## Возможности скрипта

| Режим | Что делает |
|---|---|
| `Collect` | Только чтение: автозагрузка, службы, задачи, процессы, события загрузки, питание, файл подкачки → JSON/CSV |
| `Plan` | Формирует план оптимизации по каталогу с оценкой риска и зависимостей |
| `Apply` | Применяет выбранные пункты (по умолчанию нужен `-WhatIf` для пробного прогона) с резервной копией |
| `Rollback` | Откат изменений по сохранённой резервной копии |
| `ScheduleMeasure`, `BootTrace` | Замер загрузки и трассировка WPR |
| `ListCatalog` | Список правил каталога |

Пример безопасного запуска:

```powershell
powershell -ExecutionPolicy Bypass -File .\Win11-BootRamBattery-Toolkit.ps1 -Mode Collect
powershell -ExecutionPolicy Bypass -File .\Win11-BootRamBattery-Toolkit.ps1 -Mode Apply -WhatIf
```

JSON, созданный режимом `Collect`, можно открыть в анализаторе (вкладка импорта) — данные обрабатываются только в браузере и никуда не отправляются.

## Важно

- Перед `Apply` без `-WhatIf` создайте точку восстановления.
- JSON-отчёты и журналы содержат имя компьютера и пользователя — не публикуйте их.
- `wesp` (Windows Endpoint Security Platform, KB5127936) — компонент Microsoft, не отключать.
- Страница полностью автономна: без CDN и внешних запросов.
