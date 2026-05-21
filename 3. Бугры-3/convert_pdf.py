"""Convert pptx to pdf using PowerPoint COM. Pre-normalize and use ASCII temp paths."""
import sys, os, tempfile, shutil
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client
import pythoncom
from pptx import Presentation

OUT_DIR = r"C:\Авраменко\1. КОМПАКТ\Отчеты\Еженедельные отчеты по пятницам\Бугры-3"
SRC = os.path.join(OUT_DIR, "Отчет из МСГ Бугры-3.pptx")
DST = os.path.join(OUT_DIR, "2. МСГ критические отставания СК Бугры-3 неделя 19.pdf")

NORMALIZED = os.path.join(tempfile.gettempdir(), "bugri3_normalized.pptx")
TEMP_PDF   = os.path.join(tempfile.gettempdir(), "bugri3_normalized.pdf")
Presentation(SRC).save(NORMALIZED)

pythoncom.CoInitialize()
pp = win32com.client.Dispatch("PowerPoint.Application")
pres = pp.Presentations.Open(NORMALIZED, WithWindow=False)
pres.SaveAs(TEMP_PDF, 32)
pres.Close()
pp.Quit()

shutil.copy(TEMP_PDF, DST)
print(f"PDF saved: {DST}")
print(f"Size: {os.path.getsize(DST)} bytes")
print(f"PDF saved: {DST}")
print(f"Size: {os.path.getsize(DST)} bytes")
