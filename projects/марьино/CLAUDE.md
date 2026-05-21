# CLAUDE.md — Марьино (СК «Марьино»)

Объектная специфика. Общие правила МСГ — в `МСГ\CLAUDE.md`.

**Дополнительный контекст** — `..\..\1. Марьино\README.md`: 13 корпусов (1.1–1.13) + КПП + ВУ, кураторы по разделам, статистика СМР, плюс `Марьино.md` (ручной аналитический разбор) и `Архив\` со срезами.

## Базовые факты

- **Совещание:** среда. xlsx «к совещанию» готов **во вт к концу дня**.
- **`meeting_out_dir`:** `C:\Авраменко\1. КОМПАКТ\1. Марьино\2. МСГ`.
- **gsheet:** fileId `1-xYtcP_bWVm1k8xQ325KkPnGnpBQroff77X-1Y7GwHM`, owner `gitcarazin@gmail.com`.
- **Монолит выключен** (`include_monolit: false`) — другие корпуса/отметки, matchers Репино не подходят. Когда понадобится — переписать matchers под структуру Марьино и поднять флаг.

## Quirks (обнаружено 2026-05-13)

### 1. Лист в gsheet называется `МСГ, ГПР` (с запятой и пробелом)

Скрипт автоматически находит лист по заголовку в `col 15 row 1` («Фильтр по месяцам»), но дальнейшая обработка требует переименования в `МСГ`. Само переименование критично, см. п. 3.

### 2. Excel не открывает gsheet-экспорт в обычном режиме

`Workbooks.Open(...)` → COM error `-2146827284` без причины. Открывается только с `CorruptLoad=1` (xlRepairFile). В `build_meeting_report.py` есть fallback: сначала normal, при ошибке — Repair. Что именно Excel «чинит» — не установлено (openpyxl читает файл без проблем).

### 3. Self-references `'МСГ, ГПР'!$B$2:$B$44` в формулах WORKDAY

При переименовании `МСГ, ГПР → МСГ` в обычном режиме Excel сам обновил бы ссылки. Но в Repair-режиме связь теряется → все формулы с self-ref становятся `#REF!`. Лечится в **два этапа:**

- **`strip_self_sheet_refs`** — до `Workbooks.Open` проходимся по `xl/worksheets/sheet*.xml` и убираем префикс собственного имени листа из формул (~1072 ссылки на Марьино). Функция обобщена — работает для любого имени листа из `workbook.xml`.
- **«Двойной Open»** — после `Open(CorruptLoad=1)` сразу `SaveAs(in_clean.xlsx) → Close → Open(normal)`. Без этого финальный файл сохраняется в «recovered» состоянии и при последующем открытии у пользователя Excel показывает `#REF` в формулах WORKDAY.

После этих шагов — `xl.CalculateFull()` (перекэшировать формулы перед клонированием листов).

### 4. Ожидать паттерн «Раб.ГПР с подрядчиком в X»

См. правило `HEADER_TAIL` в `build_stage_sheet` (глобальный `МСГ\CLAUDE.md`). На Марьино пока в явном виде не выловлено, но ждать — нормальная практика заполнения. Если на Пакет/Тендер/Договор пропадает Раб.ГПР — это оно.

## TODO

### Точка останова 2026-05-13 (вечер)

Пользователь правил исходник в gsheet → попросил пересобрать `/мсг марьино`. **Пересборка не завершилась:**
- Прямой URL `export?format=xlsx` отдал **HTTP 400 «Страница не найдена»** — и без cookies, и с `--browser edge`.
- MCP Google Drive `get_file_metadata` тот же fileId **видит** (modifiedTime 2026-05-13T14:29:22Z) — то есть файл существует, просто публичный export-endpoint временно не отдаёт.
- Пользователь сказал «остановись, надо уходить».

**Текущий файл:** `C:\Авраменко\1. КОМПАКТ\1. Марьино\2. МСГ\МСГ_СК Марьино отчет к совещанию от 13.05.2026.xlsx` — **СТАРАЯ** версия (до правок пользователя).

**Что попробовать (по приоритету):**

1. **Повторить прямой URL.** Google часто 400-ит сразу после save, через 5-30 минут возвращает экспорт:
   ```powershell
   python "C:\Авраменко\Claude Code Projects\МСГ\scripts\gsheet_download_xlsx.py" `
     --file-id 1-xYtcP_bWVm1k8xQ325KkPnGnpBQroff77X-1Y7GwHM `
     --out "C:\Users\amy\AppData\Local\Temp\msg_марьино_full.xlsx"
   ```

2. **Через MCP claude.ai Google Drive** — обходит публичный endpoint:
   ```
   mcp__claude_ai_Google_Drive__download_file_content(
     fileId="1-xYtcP_bWVm1k8xQ325KkPnGnpBQroff77X-1Y7GwHM",
     exportMimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
   )
   ```
   Возвращает base64 → декодировать → записать в `%TEMP%\msg_марьино_full.xlsx`. Лимит MCP 10 МБ, файл ~1.36 МБ, проходит.

3. **Дальше — стандартный пайплайн:**
   ```powershell
   Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue
   python "C:\Авраменко\Claude Code Projects\МСГ\scripts\build_meeting_report.py" `
     --project марьино --month Май `
     --xlsx "C:\Users\amy\AppData\Local\Temp\msg_марьино_full.xlsx" `
     --date <DD.MM.YYYY>
   ```
   Скрипт сам применит `strip_self_sheet_refs` + двойной Open.

### Открытые вопросы

- Подтвердить, что в текущем (старом) файле `#REF!` действительно ушёл — обратной связи от пользователя пока нет.
- Что именно правил в gsheet (формулы дат / структуру листов / праздники B2:B44).
