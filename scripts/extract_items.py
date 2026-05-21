"""Parse МСГ CSV (downloaded from Google Sheets) → items.json with section info.

Usage:
    python extract_items.py --csv <path> --out <items.json>
"""
import argparse, sys, csv, json, datetime, io

sys.stdout.reconfigure(encoding="utf-8")
csv.field_size_limit(10_000_000)

ap = argparse.ArgumentParser()
ap.add_argument("--csv", required=True)
ap.add_argument("--out", required=True)
args = ap.parse_args()

with open(args.csv, encoding="utf-8") as f:
    rows = list(csv.reader(f))

def cell(r, c):
    if r >= len(rows): return None
    if c-1 >= len(rows[r]): return None
    v = rows[r][c-1]
    return v if v != "" else None

def parse_date_or_pass(v):
    if v is None: return None
    s = str(v).strip()
    if not s: return None
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except: pass
    return s

items = []
prev_name, prev_block = None, ""
current_section, current_category = "", ""
# current_contractor — имя подрядчика, если на строке 'Раб. ГПР' сразу после
# 'Зак. ГПР' колонка X содержит значение, отличное от названия работ. Тогда X
# на этой строке — это имя подрядчика (например, «ООО "АВТОДОРКОМПЛЕКС"»).
# Если совпадает — подрядчик ещё не назначен, наследуется пустая строка.
# zak_name_for_contractor сбрасывается на следующем 'Зак. ГПР'.
current_contractor = ""
zak_name_for_contractor = None
# Раздел/Подраздел РД (D/E) — наследуются от последней непустой строки во всём листе,
# вне зависимости от наличия rtype. Нужны для матчинга «стен подвала» в Монолит-логике
# (D="Монолит ниже 0" + E="Монолитные стены").
prev_section_rd = ""
prev_subsection_rd = ""

for i in range(len(rows)):
    # Раздел/Подраздел РД — пропагация делается ДО любых continue,
    # иначе значения с пустых строк не наследуются на rtype-строки.
    d_val = cell(i, 4)
    e_val = cell(i, 5)
    if d_val: prev_section_rd = str(d_val).strip()
    if e_val: prev_subsection_rd = str(e_val).strip()

    name = cell(i, 24)
    rtype = cell(i, 38)
    block_val = cell(i, 25)

    name_s = str(name).strip() if name else None
    rtype_s = str(rtype).strip() if rtype else None
    block_s = str(block_val).strip() if block_val else None

    if name_s and not rtype_s:
        proj_obj = cell(i, 30)
        pf_nach = cell(i, 40)
        if proj_obj is None and pf_nach is None:
            current_category = name_s
            continue

    if not rtype_s:
        continue

    if rtype_s == 'Зак. ГПР' and name_s:
        current_section = name_s
        # Сбрасываем подрядчика при новом work-блоке. Имя 'Зак. ГПР' фиксируем
        # на одну итерацию — для сравнения с 'Раб. ГПР'.
        current_contractor = ""
        zak_name_for_contractor = name_s
    elif rtype_s in ('Раб. ГПР', 'Факт ГПР') and zak_name_for_contractor is not None:
        # Сырое значение X на этой строке — НЕ через наследование prev_name.
        raw_x = name
        raw_x_s = str(raw_x).strip() if raw_x else None
        if raw_x_s and raw_x_s != zak_name_for_contractor:
            current_contractor = raw_x_s
        zak_name_for_contractor = None  # больше не пытаемся ловить в этом блоке

    if name_s:
        prev_name = name_s
    # Building (Здание) is filled only on the Зак. ГПР row that starts each work group.
    # Subsequent lifecycle rows (План РД, План П, ..., План, Факт) inherit it.
    # Reset only when a new work group begins (Зак. ГПР), not on every named lifecycle row.
    if rtype_s == 'Зак. ГПР':
        prev_block = block_s if block_s else ""
    elif block_s:
        prev_block = block_s

    items.append({
        'row': i+1,
        'name': prev_name,
        'block': prev_block,
        'rtype': rtype_s,
        'section': current_section,
        'category': current_category,
        'obj_otkl': cell(i, 27),
        'proj_obj': cell(i, 30),
        'unit': cell(i, 29),
        'pct': cell(i, 33),
        'pf_val': cell(i, 32),
        'days_nach': cell(i, 35),
        'days_kon': cell(i, 36),
        'pf_nach': parse_date_or_pass(cell(i, 40)),
        'pf_okon': parse_date_or_pass(cell(i, 41)),
        'pf_month': cell(i, 43),
        'prereq': cell(i, 22),
        'section_rd':    prev_section_rd,
        'subsection_rd': prev_subsection_rd,
        'month_filter':  cell(i, 15),   # O — Фильтр по месяцам, без наследования
        'contractor':    current_contractor,
    })

with open(args.out, "w", encoding="utf-8") as f:
    json.dump(items, f, ensure_ascii=False)

print(f"Items: {len(items)}")
