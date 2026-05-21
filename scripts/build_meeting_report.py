"""Build the «отчёт к совещанию» xlsx for a project.

Input: full xlsx downloaded from gsheet by gsheet_download_xlsx.py.
Output: same xlsx, but only with the report sheets — each is a copy of the source
        МСГ-sheet with a specific AutoFilter applied.

Sheets it currently builds:
  * МСГ — col O (Фильтр по месяцам) = <month>
  * ГПР — col AL (ПЛАН/ФАКТ)        ∈ {План ГПР, Раб. ГПР, Зак. ГПР}

We use Excel COM (win32com) for the worksheet copy because openpyxl's
copy_worksheet drops conditional formatting (the entire Gantt colouring).

Usage:
    python build_meeting_report.py --project 2.репино --month Май
        --xlsx <full-gsheet-export.xlsx>
        [--date DD.MM.YYYY]
        [--out-dir <path>]
"""
import argparse, json, sys, os, datetime, shutil, time, zipfile, re, tempfile

sys.stdout.reconfigure(encoding="utf-8")


# --- pre-fix: tag «modern» Excel functions (added after 2007) with _xlfn. ---
# Google Sheets exports formulas without this prefix; the Russian-locale Excel
# fails to recognise them and demotes them to _xludf.<name> (user-defined) on
# the next save — which makes the cell show #NAME? / IFS instead of ЕСЛИМН.
MODERN_XL_FUNCS = (
    "IFS", "SWITCH", "MAXIFS", "MINIFS", "TEXTJOIN", "CONCAT", "IFNA",
    "XLOOKUP", "XMATCH", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE",
    "RANDARRAY", "LET", "LAMBDA", "MAP", "REDUCE", "BYROW", "BYCOL",
    "FORMULATEXT", "NUMBERVALUE", "ISOWEEKNUM", "ARABIC",
)
_FN_RE = re.compile(r"(?<![A-Za-z0-9_.])(" + "|".join(MODERN_XL_FUNCS) + r")\(")

def patch_modern_funcs(xlsx_path):
    """Open xlsx, prefix bare IFS/SWITCH/... with `_xlfn.` inside sheet XMLs."""
    fixed = 0
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name
    with zipfile.ZipFile(xlsx_path, "r") as src, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.namelist():
            data = src.read(item)
            if item.startswith("xl/worksheets/sheet") and item.endswith(".xml"):
                text = data.decode("utf-8")
                new_text, n = _FN_RE.subn(r"_xlfn.\1(", text)
                if n:
                    fixed += n
                    data = new_text.encode("utf-8")
            dst.writestr(item, data)
    shutil.move(tmp_path, xlsx_path)
    return fixed


def strip_self_sheet_refs(xlsx_path):
    """Убирает self-references вида 'SheetName'! из формул на том же листе.

    Зачем: gsheet Марьино называет лист «МСГ, ГПР» (с запятой) и формулы WORKDAY
    содержат self-ref `'МСГ, ГПР'!$B$2:$B$44`. Excel в обычном режиме при
    переименовании листа сам обновляет ссылки, но Марьино-xlsx открывается
    только через CorruptLoad=1 (Repair) — а Repair теряет связь имя↔ссылка,
    и переименование делает все такие формулы #REF!.

    Решение: до Workbooks.Open пройтись по xl/worksheets/sheet*.xml и убрать
    префикс собственного имени листа из формул (`'X'!Y` → `Y`, внутрилистовая
    ссылка валидна без префикса).
    """
    fixed = 0
    with zipfile.ZipFile(xlsx_path, "r") as src:
        wb_xml = src.read("xl/workbook.xml").decode("utf-8")
        try:
            rels_xml = src.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        except KeyError:
            return 0
    # sheet name → rId
    sheets = re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb_xml)
    if not sheets:
        sheets = re.findall(r'<sheet[^>]*r:id="([^"]+)"[^>]*name="([^"]+)"', wb_xml)
        sheets = [(n, r) for (r, n) in sheets]
    # rId → target
    rels = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels_xml))
    # sheet name → sheet xml path (normalize to xl/worksheets/sheetN.xml)
    name_to_path = {}
    for nm, rid in sheets:
        tgt = rels.get(rid)
        if not tgt:
            continue
        if tgt.startswith("/"):
            path = tgt.lstrip("/")
        elif tgt.startswith("xl/"):
            path = tgt
        else:
            path = "xl/" + tgt
        name_to_path[nm] = path

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name
    with zipfile.ZipFile(xlsx_path, "r") as src, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.namelist():
            data = src.read(item)
            if item.startswith("xl/worksheets/sheet") and item.endswith(".xml"):
                this_sheet = next((n for n, p in name_to_path.items() if p == item), None)
                if this_sheet:
                    text = data.decode("utf-8")
                    # Apostrophes в XML formula bodies хранятся буквально (')
                    pat_q = f"'{re.escape(this_sheet)}'!"
                    text2, n1 = re.subn(pat_q, "", text)
                    # bare (no apostrophes) — для имён без спецсимволов
                    if " " not in this_sheet and "," not in this_sheet:
                        pat_b = r"(?<![A-Za-z0-9_!])" + re.escape(this_sheet) + r"!"
                        text2, n2 = re.subn(pat_b, "", text2)
                    else:
                        n2 = 0
                    total = n1 + n2
                    if total:
                        fixed += total
                        data = text2.encode("utf-8")
            dst.writestr(item, data)
    shutil.move(tmp_path, xlsx_path)
    return fixed

# ---- args ---------------------------------------------------------------
ap = argparse.ArgumentParser()
ap.add_argument("--project", required=True)
ap.add_argument("--month", required=True)
ap.add_argument("--xlsx", required=True, help="full xlsx exported from gsheet")
ap.add_argument("--date", help="report date DD.MM.YYYY (default = today)")
ap.add_argument("--out-dir", help="override output dir")
args = ap.parse_args()

ROOT = r"C:\Авраменко\Claude Code Projects\МСГ"
cfg_path = os.path.join(ROOT, "projects", args.project, "config.json")
with open(cfg_path, encoding="utf-8") as f:
    cfg = json.load(f)

basename = cfg["report_basename"]
# out_dir приоритеты: --out-dir CLI → cfg["meeting_out_dir"] → исторический fallback для Репино.
out_dir = args.out_dir or cfg.get("meeting_out_dir") or r"C:\Авраменко\1. КОМПАКТ\6. Репино\2. МСГ"
os.makedirs(out_dir, exist_ok=True)

date_str = args.date or datetime.date.today().strftime("%d.%m.%Y")
dst_name = f"МСГ_{basename} отчет к совещанию от {date_str}.xlsx"
dst = os.path.join(out_dir, dst_name)
month = args.month.strip()

print(f"Source xlsx : {args.xlsx}")
print(f"Target xlsx : {dst}")

# ---- Excel COM SaveAs hates non-ASCII paths — work in ASCII temp ---------
TMP_DIR = r"C:\Users\amy\AppData\Local\Temp\msg_meeting"
os.makedirs(TMP_DIR, exist_ok=True)
tmp_in  = os.path.join(TMP_DIR, "in.xlsx")
tmp_out = os.path.join(TMP_DIR, "out.xlsx")
for p in (tmp_in, tmp_out):
    if os.path.exists(p):
        os.remove(p)
shutil.copyfile(args.xlsx, tmp_in)
print(f"  (copied source → {tmp_in})")

# Pre-fix: prefix «modern» functions with _xlfn. so Excel doesn't demote them
n_fixed = patch_modern_funcs(tmp_in)
print(f"  patched {n_fixed} modern-function calls with _xlfn. prefix")

# Pre-fix: убрать self-references 'SheetName'! из формул (см. docstring strip_self_sheet_refs)
n_stripped = strip_self_sheet_refs(tmp_in)
print(f"  stripped {n_stripped} self-sheet refs from formulas")

# ---- import win32com -----------------------------------------------------
import win32com.client as win32
from win32com.client import constants as c  # may be empty without makepy; we use literals

xlFilterValues = 7
xlOpenXMLWorkbook = 51

# Tab-color для «листа без минусов» — мягкий светло-зелёный (Excel «Good» style fill, RGB 198,239,206)
TAB_COLOR_OK = 198 + 239 * 256 + 206 * 65536  # BGR int = 13_561_798


def build_stage_sheet(ws, label, plan_tag, fact_tag):
    """Собрать стадийный лист (РД/Пакет/Тендер/Договор/Мобилизация).

    Правило (group-aware с подработами):
    - AutoFilter по AL ∈ {Зак. ГПР, Раб. ГПР, plan_tag, fact_tag, blank}.
    - Work-блок = строки от AL='Зак. ГПР' до следующей такой же строки.
    - Подработа внутри блока = строки от X-заполненной строки до следующей X-заполненной.
      Первая подработа = «шапка» блока (Зак.ГПР + Раб.ГПР + Факт ГПР).
    - Блок qualifying ↔ есть строка с AL ∈ ALLOWED AND AA<0.
    - В qualifying-блоке показываем: шапку + подработы, у которых AA<0.
    - Section header (AL пуст, X заполнен) виден, если в его секции есть qualifying-блок.
    - Всё остальное (не-qualifying блоки, не-qualifying подработы, пустые строки) — скрыто.

    View tweaks: outline collapsed, X..AG + V + AN..AO видны; A..W (кроме V), AH..AK, AQ скрыты.

    Требует, чтобы `xl.CalculateFull()` уже был вызван (для пересчёта формульных AA).
    """
    last_row = ws.UsedRange.Rows.Count
    flt_range = f"M1:AO{last_row}"

    if ws.AutoFilterMode:
        ws.AutoFilterMode = False
    try:
        ws.Rows.Hidden = False
    except Exception as e:
        print(f"  (warn) {label} un-hide all rows: {e}")

    tags = ("Зак. ГПР", "Раб. ГПР", plan_tag, fact_tag, "=")
    ws.Range(flt_range).AutoFilter(Field=26, Criteria1=tags, Operator=xlFilterValues)
    print(f"  {label} AutoFilter (Field 26 / AL): ∈ (Зак. ГПР, Раб. ГПР, {plan_tag}, {fact_tag}, blank)")

    data = ws.Range(ws.Cells(1, 24), ws.Cells(last_row, 38)).Value
    ALLOWED = {"Зак. ГПР", "Раб. ГПР", plan_tag, fact_tag}

    block_starts = []
    section_headers = []
    for i, row in enumerate(data):
        if i == 0:
            continue
        r = i + 1
        al = row[14]
        x  = row[0]
        x_filled  = isinstance(x, str) and x.strip()
        al_filled = isinstance(al, str) and al.strip()
        if al_filled and al.strip() == "Зак. ГПР":
            block_starts.append(r)
        elif (not al_filled) and x_filled:
            section_headers.append(r)
    block_starts.append(last_row + 1)

    blocks_with_neg = set()
    for k in range(len(block_starts) - 1):
        s, e = block_starts[k], block_starts[k + 1]
        for i in range(s - 1, e - 1):
            al = data[i][14]
            aa = data[i][3]
            if isinstance(al, str) and al.strip() in ALLOWED \
               and isinstance(aa, (int, float)) and aa < 0:
                blocks_with_neg.add((s, e))
                break

    sh_sorted = sorted(section_headers)

    visible_rows = set()
    visible_rows.add(1)
    for idx, sh in enumerate(sh_sorted):
        next_sh = sh_sorted[idx + 1] if idx + 1 < len(sh_sorted) else last_row + 1
        if any(s for (s, e) in blocks_with_neg if sh <= s < next_sh):
            visible_rows.add(sh)
    HEADER_TAIL = {"Раб. ГПР", "Факт ГПР"}
    for (s, e) in blocks_with_neg:
        subworks = []
        cur = s
        for r in range(s + 1, e):
            x = data[r - 1][0]
            al = data[r - 1][14]
            al_strip = al.strip() if isinstance(al, str) else ""
            if isinstance(x, str) and x.strip() and al_strip not in HEADER_TAIL:
                subworks.append((cur, r - 1))
                cur = r
        subworks.append((cur, e - 1))
        head = subworks[0]
        for r in range(head[0], head[1] + 1):
            visible_rows.add(r)
        for (ss, se) in subworks[1:]:
            has_neg = False
            for r in range(ss, se + 1):
                al = data[r - 1][14]
                aa = data[r - 1][3]
                if isinstance(al, str) and al.strip() in ALLOWED \
                   and isinstance(aa, (int, float)) and aa < 0:
                    has_neg = True; break
            if has_neg:
                for r in range(ss, se + 1):
                    visible_rows.add(r)

    hide_rows = [r for r in range(2, last_row + 1) if r not in visible_rows]
    ranges = []
    if hide_rows:
        s = hide_rows[0]; p = s
        for r in hide_rows[1:]:
            if r == p + 1: p = r
            else:
                ranges.append((s, p)); s = r; p = r
        ranges.append((s, p))
    for s, e in ranges:
        if s == e:
            ws.Rows(s).EntireRow.Hidden = True
        else:
            ws.Rows(f"{s}:{e}").EntireRow.Hidden = True
    print(f"  {label}: {len(block_starts)-1} блоков, {len(blocks_with_neg)} с AA<0, видимо {len(visible_rows)-1} строк (+R1), скрыто {len(hide_rows)} ({len(ranges)} диапазонов)")

    ws.Outline.ShowLevels(ColumnLevels=1)
    ws.Columns("AN:AO").EntireColumn.Hidden = False
    ws.Columns("X:AG").EntireColumn.Hidden = False
    for col_range in ("A:W", "AH:AK", "AQ:AQ"):
        ws.Columns(col_range).EntireColumn.Hidden = True
    ws.Columns("V").EntireColumn.Hidden = False  # Готовность предшественника
    print(f"  {label} view: outline collapsed, X..AG + V + AN..AO видны, A..W (кроме V) / AH..AK / AQ скрыты")

    # Tab-color: лист «без минусов» (видна только R1) → светло-зелёный — визуальный сигнал
    # «по этой стадии отставаний нет».
    if not blocks_with_neg:
        ws.Tab.Color = TAB_COLOR_OK
        print(f"  {label} tab: light green (нет минусов)")


xl = win32.gencache.EnsureDispatch("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.ScreenUpdating = False

try:
    # Маьино-экспорт gsheet выдаёт xlsx, который Excel в обычном режиме считает
    # «повреждённым» (Workbooks.Open → COM error). CorruptLoad=1 (xlRepairFile)
    # тихо чинит структуру и открывает. Делаем fallback: сначала пробуем normal,
    # потом repair — Репино/Бугры открываются normal, Марьино требует repair.
    used_repair = False
    try:
        wb = xl.Workbooks.Open(tmp_in, ReadOnly=False, UpdateLinks=0)
    except Exception as e:
        print(f"  Open normal failed → пробуем CorruptLoad=1 (repair): {e}")
        wb = xl.Workbooks.Open(tmp_in, ReadOnly=False, UpdateLinks=0, CorruptLoad=1)
        used_repair = True
        print("  Open OK через CorruptLoad=1")

    # Если открыли через Repair — пересохраняем и переоткрываем чистым, иначе
    # переименование листа ломает self-references (становятся #REF!), а Save
    # уносит «recovered» состояние в финальный файл.
    if used_repair:
        clean_path = os.path.join(TMP_DIR, "in_clean.xlsx")
        if os.path.exists(clean_path):
            os.remove(clean_path)
        wb.SaveAs(clean_path, FileFormat=xlOpenXMLWorkbook)
        wb.Close(SaveChanges=False)
        print(f"  re-saved (post-repair) → {clean_path}, reopening normally")
        wb = xl.Workbooks.Open(clean_path, ReadOnly=False, UpdateLinks=0)
        print("  Open OK после resave")

    # Calculate сразу после Open, чтобы все формулы получили актуальные значения
    # и закэшировались. Особенно важно для WORKDAY со ссылками на праздники.
    xl.CalculateFull()
    print("  CalculateFull (post-open)")

    # Find МСГ sheet by header (col 15 row 1 contains «Фильтр по месяцам»).
    msg_src = None
    for ws in wb.Worksheets:
        v = ws.Cells(1, 15).Value
        if isinstance(v, str) and "фильтр по месяц" in v.lower():
            msg_src = ws
            break
    if msg_src is None:
        names = [s.Name for s in wb.Worksheets]
        raise SystemExit(f"МСГ sheet not found. Sheets: {names}")
    print(f"  МСГ sheet found: {msg_src.Name!r}")

    # Drop every other sheet first (РС etc.). Workbook must keep ≥1 sheet, so
    # the source has to survive — keep it and delete the rest.
    src_name = msg_src.Name
    for ws in list(wb.Worksheets):
        if ws.Name != src_name:
            print(f"  removing sheet: {ws.Name!r}")
            ws.Delete()

    # Now we have exactly one sheet = src. We need TWO copies of it:
    # one for МСГ (filter O=<month>) and one for ГПР (filter AL=stage tags).
    # Strategy: rename src → 'МСГ', then Copy(After=src) → that becomes 'ГПР'.
    src = wb.Worksheets(src_name)

    # Clear any inherited AutoFilter / hidden rows from gsheet so the copy starts clean
    if src.AutoFilterMode:
        src.AutoFilterMode = False
    # Un-hide every row on the source (we'll let AutoFilter re-hide based on the criteria)
    try:
        src.Rows.Hidden = False
    except Exception as e:
        print(f"  (warn) couldn't un-hide rows globally: {e}")

    src.Name = "МСГ"
    print(f"  renamed source → 'МСГ'")

    # Excel.Worksheet.Copy(After=…) duplicates the sheet WITH conditional formatting,
    # column widths, merges, formulas — everything.
    wb.Worksheets("МСГ").Copy(After=wb.Worksheets("МСГ"))
    # The new sheet is auto-named "МСГ (2)"
    gpr = wb.Worksheets(wb.Worksheets.Count)
    gpr.Name = "ГПР"
    print(f"  cloned МСГ → 'ГПР'")

    # Clone once more for РД (sheet order: МСГ, ГПР, РД)
    wb.Worksheets("МСГ").Copy(After=wb.Worksheets("ГПР"))
    rd = wb.Worksheets(wb.Worksheets.Count)
    rd.Name = "РД"
    print(f"  cloned МСГ → 'РД'")

    # Clone once more for Пакет (sheet order: МСГ, ГПР, РД, Пакет)
    wb.Worksheets("МСГ").Copy(After=wb.Worksheets("РД"))
    packet = wb.Worksheets(wb.Worksheets.Count)
    packet.Name = "Пакет"
    print(f"  cloned МСГ → 'Пакет'")

    # Clone once more for Тендер (sheet order: МСГ, ГПР, РД, Пакет, Тендер)
    wb.Worksheets("МСГ").Copy(After=wb.Worksheets("Пакет"))
    tender = wb.Worksheets(wb.Worksheets.Count)
    tender.Name = "Тендер"
    print(f"  cloned МСГ → 'Тендер'")

    # Clone once more for Договор (sheet order: …, Тендер, Договор)
    wb.Worksheets("МСГ").Copy(After=wb.Worksheets("Тендер"))
    dogovor = wb.Worksheets(wb.Worksheets.Count)
    dogovor.Name = "Договор"
    print(f"  cloned МСГ → 'Договор'")

    # Clone once more for Мобилизация (sheet order: …, Договор, Мобилизация)
    wb.Worksheets("МСГ").Copy(After=wb.Worksheets("Договор"))
    mobil = wb.Worksheets(wb.Worksheets.Count)
    mobil.Name = "Мобилизация"
    print(f"  cloned МСГ → 'Мобилизация'")

    # --- Apply filters ----------------------------------------------------
    # Range "M1:AO<lastRow>" is the same filter range gsheet uses.
    msg = wb.Worksheets("МСГ")
    gpr = wb.Worksheets("ГПР")
    rd  = wb.Worksheets("РД")
    packet = wb.Worksheets("Пакет")
    tender = wb.Worksheets("Тендер")
    dogovor = wb.Worksheets("Договор")
    mobil = wb.Worksheets("Мобилизация")
    last_row = msg.UsedRange.Rows.Count
    flt_range = f"M1:AO{last_row}"

    # МСГ: O is column 15 = field 3 in M..AO range (M=1, N=2, O=3)
    msg.Range(flt_range).AutoFilter(Field=3, Criteria1=month)
    print(f"  МСГ AutoFilter: O = {month!r}")

    # МСГ view tweaks:
    #  - collapse the Gantt outline to month level
    #  - keep all columns X..AO visible (un-hide them after outline collapse)
    #  - re-show the reporting month columns in the Gantt area
    #  - hide A:W
    #  - rebuild conditional formatting on AI:AJ: <0 → red fill, >0 → green fill
    msg.Outline.ShowLevels(ColumnLevels=1)
    print("  МСГ Gantt outline collapsed to month level")

    msg.Columns("X:AO").EntireColumn.Hidden = False
    print("  МСГ un-hidden: X..AO (everything between Наименование и ПЛАН/ФАКТ окончание)")

    # Re-expand the reporting month columns. Lookup is dynamic: parse the year
    # from --date, map Russian month to a number, then scan header row for date
    # cells in that year+month.
    RU_MONTHS = {"январь":1,"февраль":2,"март":3,"апрель":4,"май":5,"июнь":6,
                 "июль":7,"август":8,"сентябрь":9,"октябрь":10,"ноябрь":11,"декабрь":12}
    month_num = RU_MONTHS.get(month.strip().lower())
    try:
        year_num = int(date_str.split(".")[-1])
    except Exception:
        year_num = datetime.date.today().year
    expanded = 0
    if month_num:
        max_col = msg.UsedRange.Columns.Count
        for cc in range(44, max_col + 1):
            v = msg.Cells(1, cc).Value
            if hasattr(v, "year") and v.year == year_num and v.month == month_num:
                msg.Columns(cc).EntireColumn.Hidden = False
                expanded += 1
    print(f"  МСГ re-expanded {expanded} columns for {month} {year_num}")

    msg.Columns("A:W").EntireColumn.Hidden = True
    msg.Columns("V").EntireColumn.Hidden = False  # «Готовность предшественника» оставляем видимой
    print("  МСГ hidden cols: A..W (except V — Готовность предшественника)")

    # AI:AJ (rows 2+) — replace existing conditional formatting with sign-based
    # font colours only (no fill, so any manual fill stays visible):
    #   <0 → red font
    #   >0 → green font
    # Header row 1 is excluded from the range so its formatting is untouched.
    xlCellValue = 1
    xlLess = 6
    xlGreater = 5
    last_row = msg.UsedRange.Rows.Count
    try:
        rng = msg.Range(f"AI2:AJ{last_row}")
        rng.FormatConditions.Delete()
        neg = rng.FormatConditions.Add(Type=xlCellValue, Operator=xlLess, Formula1="0")
        neg.Font.Color = 255          # RGB(255, 0, 0) — red
        neg.StopIfTrue = False
        pos = rng.FormatConditions.Add(Type=xlCellValue, Operator=xlGreater, Formula1="0")
        pos.Font.Color = 32768        # RGB(0, 128, 0) — dark green
        pos.StopIfTrue = False
        print(f"  МСГ AI:AJ CF set (rows 2..{last_row}): <0 red font, >0 green font")
    except Exception as e:
        print(f"  МСГ AI:AJ CF setup warning: {e}")

    # ГПР: AL is column 38 = field 26 in M..AO range
    # filter values + blanks → pass "=" as one of the values (Excel idiom)
    gpr_tags = ("Зак. ГПР", "Раб. ГПР", "Факт ГПР", "=")
    gpr.Range(flt_range).AutoFilter(Field=26, Criteria1=gpr_tags, Operator=xlFilterValues)
    print(f"  ГПР AutoFilter: AL ∈ (Зак.ГПР, Раб.ГПР, Факт ГПР, blank)")

    # ГПР: collapse the right-side Gantt outline to month level first (level 1 = month
    # separators visible, individual days hidden). gsheet groups also include AN/AO/AP/AQ
    # at outline=1 — we'll re-show AN..AO below.
    gpr.Outline.ShowLevels(ColumnLevels=1)
    print("  ГПР Gantt outline collapsed to month level")

    # Re-show AN («ПЛАН/ФАКТ начало») and AO («ПЛАН/ФАКТ окончание») even though they
    # were swept up into the month-level outline group.
    gpr.Columns("AN:AO").EntireColumn.Hidden = False
    print("  ГПР un-hidden: AN:AO (план/факт начало/окончание)")

    # ГПР: hide columns A..W, Z..AK, AQ (leave X = «Наименование работ», Y = «Здание»,
    # AL = ПЛАН/ФАКТ, AN/AO = даты, AR..end = Gantt month separators).
    for col_range in ("A:W", "Z:AK", "AQ:AQ"):
        gpr.Columns(col_range).EntireColumn.Hidden = True
    # AG («ОБЪЁМ ПРОЕКТНЫЙ») — оставляем видимым на ГПР для всех объектов
    # (пользовательский запрос 2026-05-13).
    gpr.Columns("AG").EntireColumn.Hidden = False
    print("  ГПР hidden cols: A..W, Z..AK, AQ (AG оставлен видимым)")

    # --- Стадийные листы (group-aware с подработами) ---
    # Все 5 листов используют один и тот же алгоритм (см. build_stage_sheet),
    # отличаются только AL-тегами стадии.
    # CalculateFull нужен один раз перед всеми листами — чтобы формульные AA (=IF(...))
    # выдавали число, по которому проверяем «< 0».
    xl.CalculateFull()

    build_stage_sheet(rd,      "РД",          "План РД", "Факт РД")

    build_stage_sheet(packet,  "Пакет",       "План П",  "Факт П")

    build_stage_sheet(tender,  "Тендер",      "План Т",  "Факт Т")

    build_stage_sheet(dogovor, "Договор",     "План Д",  "Факт Д")

    build_stage_sheet(mobil,   "Мобилизация", "План М",  "Факт М")

    if cfg.get("include_monolit", True):
        # --- Build the Монолит sheet ------------------------------------------
        # Fixed 13-row template (matches «форматирование таблицы.jpg» reference):
        #   13 work types (fund + 6 wall levels + 6 slab levels) × 6 buildings × {План, Факт}.
        # For each template row we collect data from МСГ by matching name patterns.
        BUILDINGS = ("К1","К2","К3","К4","К5","К6")
        import re as _re

        # (label, matcher) — matcher takes (x_name, d_section, e_subsection) and returns bool
        def _has(text, *needles):
            if not isinstance(text, str): return False
            s = text.lower()
            return all((n.lower() in s) for n in needles)
        def _any(text, *needles):
            if not isinstance(text, str): return False
            s = text.lower()
            return any((n.lower() in s) for n in needles)

        # Matchers — clean rules:
        #   стены — по «N-ого этажа»  (отметка не используется, она может быть многозначной)
        #   перекрытия — по номеру ПП<N>  (или D=Монолит ниже 0 + E=Монолитные перекрытия для подвала)
        #   фундамент — по «фундамент»
        #   стены подвала — D=Монолит ниже 0 + E=Монолитные стены
        TEMPLATE = [
            ("ФУНДАМЕНТНАЯ ПЛИТА",
             lambda x, d, e: _has(x, "фундамент")),
            ("МОНОЛИТНЫЕ Ж/Б СТЕНЫ НИЖЕ ОТМ. 0.000 (-1 эт.)",
             lambda x, d, e: ("Монолит ниже 0" in (d or "")) and ("Монолитные стены" in (e or ""))),
            ("МОНОЛИТНЫЕ Ж/Б ПЕРЕКРЫТИЯ  НА ОТМ. -0.100 (1 эт.)",
             lambda x, d, e: _has(x, "перекрыти") and "-0.100" in (x or "")),
            ("МОНОЛИТНЫЕ Ж/Б СТЕНЫ НИЖЕ ОТМ. +3.200 (1 эт.)",
             lambda x, d, e: _has(x, "стен", "1-ого этажа")),
            ("МОНОЛИТНЫЕ Ж/Б ПЕРЕКРЫТИЯ  НА ОТМ. +3.200 (2 эт.)",
             lambda x, d, e: _has(x, "перекрыти", "ПП1")),
            ("МОНОЛИТНЫЕ Ж/Б СТЕНЫ НИЖЕ ОТМ. +6.150 (2 эт.)",
             lambda x, d, e: _has(x, "стен", "2-ого этажа")),
            ("МОНОЛИТНЫЕ Ж/Б ПЕРЕКРЫТИЯ  НА ОТМ. +6.150 (3 эт.)",
             lambda x, d, e: _has(x, "перекрыти", "ПП2")),
            ("МОНОЛИТНЫЕ Ж/Б СТЕНЫ НИЖЕ ОТМ. +9.800 (3 эт.)",
             lambda x, d, e: _has(x, "стен", "3-ого этажа")),
            ("МОНОЛИТНЫЕ Ж/Б ПЕРЕКРЫТИЯ  НА ОТМ. +9.800 (4 эт.)",
             lambda x, d, e: _has(x, "перекрыти", "ПП3")),
            ("МОНОЛИТНЫЕ Ж/Б СТЕНЫ НИЖЕ ОТМ. +13.400 (4 эт.)",
             lambda x, d, e: _has(x, "стен", "4-ого этажа")),
            ("МОНОЛИТНЫЕ Ж/Б ПЕРЕКРЫТИЯ  НА ОТМ. +13.400 (5 эт.)",
             lambda x, d, e: _has(x, "перекрыти", "ПП4")),
            ("МОНОЛИТНЫЕ Ж/Б СТЕНЫ НИЖЕ ОТМ. +16.700 (5 эт.)",
             lambda x, d, e: _has(x, "стен", "5-ого этажа")),
            ("МОНОЛИТНЫЕ Ж/Б ПЕРЕКРЫТИЯ  НА  ОТМ. +16.700",
             lambda x, d, e: _has(x, "перекрыти", "ПП5")),
        ]

        # Scan МСГ
        print(f"  Монолит: scanning МСГ for бетонирование rows in {month}…")
        msg_last_row = msg.UsedRange.Rows.Count
        # table[row_idx][building] → {"План": sum, "Факт": sum}
        table = [{b: {"План": 0.0, "Факт": 0.0} for b in BUILDINGS} for _ in TEMPLATE]
        prev_x, prev_y, prev_d, prev_e = None, None, None, None
        matched_count = 0
        for r in range(2, msg_last_row + 1):
            x  = msg.Cells(r, 24).Value   # Наименование
            y  = msg.Cells(r, 25).Value   # Здание
            d  = msg.Cells(r,  4).Value   # Раздел РД
            e  = msg.Cells(r,  5).Value   # Подраздел РД
            al = msg.Cells(r, 38).Value   # ПЛАН/ФАКТ
            aq = msg.Cells(r, 43).Value   # ПЛАН/ФАКТ МЕСЯЦ
            of = msg.Cells(r, 15).Value   # Фильтр по месяцам
            # propagate previous values for inherited cells
            if x: prev_x = x
            else: x = prev_x
            if y: prev_y = y
            else: y = prev_y
            if d: prev_d = d
            else: d = prev_d
            if e: prev_e = e
            else: e = prev_e
            if not isinstance(x, str) or not x.startswith("Бетонирование"): continue
            if y not in BUILDINGS: continue
            if al not in ("План", "Факт"): continue
            if aq is None: continue
            if not isinstance(of, str) or of.strip().lower() != month.strip().lower(): continue
            try:
                v = float(aq)
            except (TypeError, ValueError):
                continue
            # match to template row (first match wins)
            for idx, (_lbl, matcher) in enumerate(TEMPLATE):
                try:
                    if matcher(x, d, e):
                        table[idx][y][al] += v
                        matched_count += 1
                        break
                except Exception:
                    pass
        print(f"  Монолит: matched {matched_count} МСГ rows to {len(TEMPLATE)} template rows")

        # Add new sheet
        mon = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        mon.Name = "Монолит"

        # Headers (rows 2..4)
        mon.Range("B2").Value = "Вид монолита"
        mon.Range("C2").Value = "Месяц"
        mon.Range("D2").Value = "№ корпуса"
        mon.Range("B2:B4").Merge()
        mon.Range("C2:C4").Merge()
        mon.Range("D2:O2").Merge()
        for i, b in enumerate(BUILDINGS):
            col_start = 4 + i * 2
            mon.Cells(3, col_start).Value = i + 1
            mon.Range(mon.Cells(3, col_start), mon.Cells(3, col_start + 1)).Merge()
            mon.Cells(4, col_start).Value     = "План, м3"
            mon.Cells(4, col_start + 1).Value = "Факт, м3"

        # Data rows (one per TEMPLATE row, always all 13)
        # Values rounded to integers per user request 2026-05-12.
        for i, (label, _matcher) in enumerate(TEMPLATE):
            r = 5 + i
            mon.Cells(r, 2).Value = label
            mon.Cells(r, 3).Value = month                # «Май» as plain text (matches the screenshot)
            for j, b in enumerate(BUILDINGS):
                cs = 4 + j * 2
                d = table[i][b]
                if d["План"]:
                    mon.Cells(r, cs).Value = int(round(d["План"]))
                if d["Факт"]:
                    mon.Cells(r, cs + 1).Value = int(round(d["Факт"]))

        # ИТОГО
        total_row = 5 + len(TEMPLATE)
        mon.Cells(total_row, 2).Value = "ИТОГО:"
        for j in range(6):
            cs = 4 + j * 2
            for off in (0, 1):
                col = cs + off
                col_letter = ""
                c2 = col
                while c2 > 0:
                    col_letter = chr((c2 - 1) % 26 + ord("A")) + col_letter
                    c2 = (c2 - 1) // 26
                mon.Cells(total_row, col).Formula = f"=SUM({col_letter}5:{col_letter}{total_row-1})"

        # --- Formatting — 1-to-1 with reference 'Монолит ' from snapshot 11.05.2026 ---
        # In the source, fills use theme=0 (background 1 = white) with tints:
        #   tint = -0.15  → BG 1 darker 15%  = #D9D9D9 = RGB(217,217,217) = BGR int 14277081
        #   tint = -0.05  → BG 1 darker  5%  = #F2F2F2 = RGB(242,242,242) = BGR int 15921906
        # План columns (D, F, H, J, L, N)  → DARK
        # Факт columns (E, G, I, K, M, O)  → LIGHT
        # B (Вид) and C (Месяц) columns    → LIGHT in data rows + ИТОГО, DARK in header rows 2..3
        # Header row 4: «План, м3» = DARK, «Факт, м3» = LIGHT
        DARK  = 14277081  # #D9D9D9
        LIGHT = 15921906  # #F2F2F2

        # Whole table — Times New Roman
        tbl = mon.Range(f"B2:O{total_row}")
        tbl.Font.Name = "Times New Roman"

        # Header rows 2..3 — entirely DARK + bold + center
        hdr_top = mon.Range("B2:O3")
        hdr_top.Interior.Color = DARK
        hdr_top.Font.Bold = True
        hdr_top.HorizontalAlignment = -4108
        hdr_top.VerticalAlignment   = -4108
        hdr_top.WrapText = True

        # Header row 4 — alternating: B,C unused; План=DARK; Факт=LIGHT
        for col_idx in (4, 6, 8, 10, 12, 14):  # D F H J L N — План
            mon.Cells(4, col_idx).Interior.Color = DARK
        for col_idx in (5, 7, 9, 11, 13, 15):  # E G I K M O — Факт
            mon.Cells(4, col_idx).Interior.Color = LIGHT
        mon.Range("B4:C4").Interior.Color = DARK  # carry header style for B/C
        mon.Range("B4:O4").Font.Bold = True
        mon.Range("B4:O4").HorizontalAlignment = -4108
        mon.Range("B4:O4").VerticalAlignment   = -4108

        # Data rows 5..(total_row-1)
        data_first = 5
        data_last  = total_row - 1
        # B (label) + C (month) — LIGHT
        mon.Range(f"B{data_first}:C{data_last}").Interior.Color = LIGHT
        # План columns — DARK
        for col_letter in ("D","F","H","J","L","N"):
            mon.Range(f"{col_letter}{data_first}:{col_letter}{data_last}").Interior.Color = DARK
        # Факт columns — LIGHT
        for col_letter in ("E","G","I","K","M","O"):
            mon.Range(f"{col_letter}{data_first}:{col_letter}{data_last}").Interior.Color = LIGHT
        # Center «Месяц» values
        mon.Range(f"C{data_first}:C{data_last}").HorizontalAlignment = -4108

        # ИТОГО row — same alternating scheme + bold
        mon.Range(f"B{total_row}:C{total_row}").Interior.Color = LIGHT
        for col_letter in ("D","F","H","J","L","N"):
            mon.Range(f"{col_letter}{total_row}").Interior.Color = DARK
        for col_letter in ("E","G","I","K","M","O"):
            mon.Range(f"{col_letter}{total_row}").Interior.Color = LIGHT
        total_rng = mon.Range(f"B{total_row}:O{total_row}")
        total_rng.Font.Bold = True
        total_rng.HorizontalAlignment = -4152  # xlRight

        # Borders for whole table (thin all-around)
        for b_idx in (7, 8, 9, 10, 11, 12):
            try:
                tbl.Borders(b_idx).LineStyle = 1
                tbl.Borders(b_idx).Weight    = 2  # xlThin
            except Exception:
                pass

        # Alignment & number format for data area
        data_rng = mon.Range(f"D5:O{total_row}")
        data_rng.HorizontalAlignment = -4152  # xlRight
        try:
            data_rng.NumberFormat = "0.00"
        except Exception as e:
            print(f"  Монолит NumberFormat warning: {e}")

        # Column widths
        mon.Columns("B").ColumnWidth = 45
        mon.Columns("C").ColumnWidth = 8
        for L in ("D","E","F","G","H","I","J","K","L","M","N","O"):
            mon.Columns(L).ColumnWidth = 10
        # Row heights
        mon.Rows("2:4").RowHeight = 22
        for rr in range(5, total_row + 1):
            mon.Rows(rr).RowHeight = 18
        print(f"  Монолит: записан ({len(TEMPLATE)} строк), ИТОГО на R{total_row}")

    else:
        print("  Монолит: пропущен (include_monolit=false в config)")

    # Activate МСГ so it's the open sheet on next launch
    msg.Activate()
    msg.Range("A1").Select()

    # --- Save -------------------------------------------------------------
    if os.path.exists(tmp_out):
        os.remove(tmp_out)
    wb.SaveAs(tmp_out, FileFormat=xlOpenXMLWorkbook)
    wb.Close(SaveChanges=False)
    print(f"  saved → {tmp_out}")
finally:
    xl.Quit()

# Copy back to the (Cyrillic) target path
if os.path.exists(dst):
    os.remove(dst)
shutil.copyfile(tmp_out, dst)
print(f"Done: {dst}")
