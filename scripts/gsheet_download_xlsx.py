"""Download a Google Sheet as full XLSX (with formatting, formulas, conditional fmt).

This bypasses the 10 MB cap of the MCP `download_file_content` tool by going
straight to the public export URL:

    https://docs.google.com/spreadsheets/d/<id>/export?format=xlsx

This works as long as the gsheet is shared «anyone with the link» (which is
the case for our МСГ_RBI files). If access is restricted, pass --browser to
attach the user's google cookies via browser_cookie3.

Usage:
    python gsheet_download_xlsx.py --file-id <ID> --out <path.xlsx> [--browser chrome|edge|firefox]
"""
import argparse, sys, os

sys.stdout.reconfigure(encoding="utf-8")

ap = argparse.ArgumentParser()
ap.add_argument("--file-id", required=True)
ap.add_argument("--out", required=True)
ap.add_argument("--browser", default=None, choices=["chrome", "edge", "firefox", "brave"],
                help="optional: attach google cookies from this browser (for restricted files)")
args = ap.parse_args()

import requests

url = f"https://docs.google.com/spreadsheets/d/{args.file_id}/export?format=xlsx"
print(f"GET {url}")

session = requests.Session()
if args.browser:
    import browser_cookie3
    loader = getattr(browser_cookie3, args.browser)
    session.cookies = loader(domain_name="google.com")
    print(f"Cookies attached from {args.browser}")

r = session.get(url, allow_redirects=True, timeout=300)
print(f"Status: {r.status_code}, content-type: {r.headers.get('Content-Type')}, size: {len(r.content)} bytes")

if r.status_code != 200:
    sys.exit(f"Failed: HTTP {r.status_code}\nFirst 500 bytes:\n{r.content[:500]!r}")
ct = r.headers.get("Content-Type", "")
if "spreadsheetml" not in ct and "officedocument" not in ct:
    sys.exit(f"Wrong content-type (file may be restricted): {ct}\n"
             "Try --browser chrome|edge|firefox to attach your cookies.\n"
             f"First 500 bytes:\n{r.content[:500]!r}")

os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
with open(args.out, "wb") as f:
    f.write(r.content)
print(f"Saved: {args.out}")
