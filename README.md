# Win11 Optimizer — загрузка, ОЗУ и батарея (Windows 11 Pro 26H2)

Автономный HTML-анализатор и PowerShell-программа для оценки влияния автозагрузки, служб и фоновых процессов на время загрузки, расход оперативной памяти и время работы от батареи.

- **Онлайн-анализатор:** https://wadi-li.github.io/win11-optimizer/
- **Скрипт:** [`Win11-BootRamBattery-Toolkit.ps1`](Win11-BootRamBattery-Toolkit.ps1) (v2.0.0, PowerShell 5.1+)

## Возможности скрипта

| Режим | Что делает |
|---|---|
| `Menu` | Интерактивное меню (по умолчанию) |
| `Collect` | Только чтение: автозагрузка, службы, задачи, процессы, события загрузки, питание, файл подкачки → JSON/CSV |
| `Plan` | План оптимизации по каталогу с уровнем риска (`-Level None/Low/Medium`, `-Include`, `-Exclude`) |
| `Apply` | Применение плана: точка восстановления и файл отката; с `-WhatIf` — пробный прогон без изменений |
| `Rollback` | Откат по сохранённому файлу (`-RollbackFile`) |
| `ScheduleMeasure` / `UnscheduleMeasure` | Задача замера времени загрузки после входа |
| `BootTraceStart` / `BootTraceStop` / `BootTraceCancel` | Трассировка загрузки через WPR |
| `ListCatalog` | Список правил каталога |

Пример безопасного запуска:

```powershell
powershell -ExecutionPolicy Bypass -File .\Win11-BootRamBattery-Toolkit.ps1 -Mode Collect
powershell -ExecutionPolicy Bypass -File .\Win11-BootRamBattery-Toolkit.ps1 -Mode Apply -WhatIf
```

JSON, созданный режимом `Collect`, можно открыть в анализаторе (вкладка импорта) — данные обрабатываются только в браузере и никуда не отправляются.

## Важно

- Сначала запускайте `Apply` с `-WhatIf`; при реальном применении скрипт сам создаёт точку восстановления (отключается `-NoRestorePoint`).
- JSON-отчёты и журналы содержат имя компьютера и пользователя — не публикуйте их.
- `wesp` (Windows Endpoint Security Platform, KB5127936) — компонент Microsoft, не отключать.
- Страница полностью автономна: без CDN и внешних запросов.
