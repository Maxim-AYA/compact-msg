"""Convert .pptx → .pdf via PowerPoint COM. Pre-normalizes XML through python-pptx,
saves PDF to ASCII temp path then copies to final destination (PowerPoint COM
chokes on Cyrillic SaveAs paths).

Usage:
    python convert_pdf.py --pptx <path> --pdf <path>
"""
import argparse, sys, os, tempfile, shutil
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client, pythoncom
from pptx import Presentation

ap = argparse.ArgumentParser()
ap.add_argument("--pptx", required=True)
ap.add_argument("--pdf", required=True)
args = ap.parse_args()

NORMALIZED = os.path.join(tempfile.gettempdir(), "msg_report_normalized.pptx")
TEMP_PDF   = os.path.join(tempfile.gettempdir(), "msg_report_normalized.pdf")
Presentation(args.pptx).save(NORMALIZED)

pythoncom.CoInitialize()
pp = win32com.client.Dispatch("PowerPoint.Application")
pres = pp.Presentations.Open(NORMALIZED, WithWindow=False)
pres.SaveAs(TEMP_PDF, 32)  # 32 = ppSaveAsPDF
pres.Close()
pp.Quit()

os.makedirs(os.path.dirname(args.pdf), exist_ok=True)
shutil.copy(TEMP_PDF, args.pdf)
print(f"PDF saved: {args.pdf}")
print(f"Size: {os.path.getsize(args.pdf)} bytes")
