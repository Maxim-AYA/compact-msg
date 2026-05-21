import openpyxl
from pathlib import Path

p = Path(r"C:\Авраменко\Claude Code Projects\МСГ\2. Репино\МСГ_RBI Репино Санаторий.xlsx")
wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
ws = wb['МСГ, ГПР']

# Find the stage column. From analysis: each work has 14 lifecycle rows.
# Likely stage marker is in a column. Let's scan rows 50..200 cols 1..43 for short labels.
print("--- Rows 50..120, key columns 17 (Ответственный), 18 (ТЭГ), 24 (Наименование), 25 (Здание), 38 (План/Факт), 40, 41 ---")
for ri, row in enumerate(ws.iter_rows(min_row=50, max_row=120, values_only=True), start=50):
    c17 = row[16]; c18 = row[17]; c20 = row[19]; c24 = row[23]; c25 = row[24]
    c38 = row[37]; c40 = row[39] if len(row)>39 else None; c41 = row[40] if len(row)>40 else None
    c33 = row[32]
    has_any = any(v is not None for v in [c17,c18,c20,c24,c25,c38,c40,c41,c33])
    if has_any:
        print(f"r{ri}: resp={str(c17)[:14]!r} | tag={str(c18)[:10]!r} | id={str(c20)[:6]!r} | name={str(c24)[:30]!r} | bld={str(c25)[:6]!r} | pf={str(c38)[:10]!r} | %={str(c33)[:8]!r} | {str(c40)[:10]} - {str(c41)[:10]}")
