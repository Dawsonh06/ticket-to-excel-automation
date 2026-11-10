"""
test_2610_qr.py — offline checks for new-format (2610) QR routing.

    python test_2610_qr.py TRACKER_DIR [OLD_TRACKER ...]

TRACKER_DIR holds ticket_tracker_{Permanent Material,Dump Out,Trucking}.xlsx
(build them with setup_2610_trackers.py).  Any OLD_TRACKER files given
(e.g. ticket_tracker_2602.xlsx) have every old-format QR code in them
re-parsed to prove old-format parsing is unchanged.  No network calls.
"""
import os
import re
import shutil
import sys
import tempfile

import openpyxl

import ticket_processor as tp

# The old-format pattern and parser exactly as they were before 2610 support.
_OLD_PATTERN = re.compile(
    r'\b([A-Z0-9]{4})-([A-Z0-9]{4})-([A-Z0-9]{6})-([A-Z0-9]{2})\b', re.IGNORECASE)


def old_parse(text):
    m = _OLD_PATTERN.search(text)
    if not m:
        return None
    a, b, c, d = (g.upper() for g in m.groups())
    return {"job_number": a, "location": b, "cost_code_1": c, "cost_code_2": d,
            "cost_code": f"{c}-{d}", "raw": text}


def route(qr_text, tracker_dir):
    qr = tp._parse_any_qr(qr_text)
    if qr is None:
        return "no QR pattern matched"
    if qr.get("format") != "new":
        same = qr == old_parse(qr_text)
        return (f"old format → existing path (job {qr['job_number']}, cost code "
                f"{qr['cost_code']}); identical to pre-change parse: {same}")
    r = tp._nf_route(qr, lambda k: os.path.join(
        tracker_dir, tp._nf_tracker_filename(qr["job_number"], k)))
    if not r["ok"]:
        return f"FLAGGED: {r['reason']}"
    return f"{os.path.basename(r['excel_path'])}, tab {r['toc_row']['tab']!r}"


CASES = [
    ("01-2610-8990-999999-01", "rock tracker, tab 10"),
    ("02-2610-8990-999999-02", "dump tracker, tab 14"),
    ("03-2610-8990-999999-03", "trucking tracker, tab 18"),
    ("01-2610-1111-111111",    "rock tracker, Other Work"),
    ("03-2610-1111-111111",    "trucking tracker, Other Work"),
    ("2601-0180-313713-99",    "unchanged 2601 behavior"),
    ("04-2610-8990-999999-01", "flagged, invalid type prefix"),
    ("01-2610-8990-999999-09", "flagged, not in TOC"),
]
EXPECT = [
    "ticket_tracker_Permanent Material.xlsx, tab '10'",
    "ticket_tracker_Dump Out.xlsx, tab '14'",
    "ticket_tracker_Trucking.xlsx, tab '18'",
    "ticket_tracker_Permanent Material.xlsx, tab 'Other Work'",
    "ticket_tracker_Trucking.xlsx, tab 'Other Work'",
    "old format",
    "FLAGGED: Invalid type prefix",
    "FLAGGED: QR 01-2610-8990-999999-09 not found in 2610 TOC",
]


def main():
    tracker_dir = sys.argv[1]
    failures = 0

    print("\n=== Routing: the 8 required cases ===")
    for (qr, want), expect in zip(CASES, EXPECT):
        got = route(qr, tracker_dir)
        ok  = got.startswith(expect) and ("True" in got if expect == "old format" else True)
        failures += not ok
        print(f"{'PASS' if ok else 'FAIL'}  {qr:<24} expected: {want:<30} got: {got}")

    print("\n=== New-format codes inside surrounding text (OCR / sticker text) ===")
    for text in ("QR: 01-2610-8990-999999-01 Storm", "x 03-2610-1111-111111\n"):
        qr = tp._parse_any_qr(text, whole_text_is_code=False)
        ok = qr and qr.get("format") == "new"
        failures += not ok
        print(f"{'PASS' if ok else 'FAIL'}  {text!r} → {qr and qr['raw']}")
    # The old pattern alone would have mis-read the tail as an old 2610 code:
    print(f"      (old pattern on '01-2610-8990-999999-01' would have matched "
          f"{_OLD_PATTERN.search('01-2610-8990-999999-01').group(0)!r}; "
          f"new QR_PATTERN matches: {tp.QR_PATTERN.search('01-2610-8990-999999-01')})")

    print("\n=== Regression: every old-format QR in the 2601/2602 trackers ===")
    for path in sys.argv[2:]:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        codes = set()
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    if v is not None and _OLD_PATTERN.search(str(v)):
                        codes.add(str(v).strip())
        wb.close()
        bad = [c for c in codes if tp._parse_any_qr(c) != old_parse(c)
               or tp._parse_any_qr(c, whole_text_is_code=False)
               != tp._parse_qr_string(_OLD_PATTERN.search(c).group(0))]
        failures += bool(bad)
        print(f"{'PASS' if not bad else 'FAIL'}  {os.path.basename(path)}: "
              f"{len(codes)} QR value(s) parse identically" + (f" — differ: {bad}" if bad else ""))

    print("\n=== Write test (temp copies; nothing uploaded) ===")
    tmp = tempfile.mkdtemp()
    try:
        dump = shutil.copy(os.path.join(tracker_dir, "ticket_tracker_Dump Out.xlsx"), tmp)
        r = tp._nf_route(tp._parse_any_qr("02-2610-8990-999999-02"), lambda k: dump)
        # Dump tab 14 expects CY; this ticket only shows tons → yellow + note
        tp.write_to_excel_new_format(dump, "dump", r["toc_row"], {
            "date": "11/10/2026", "company": "Hillsboro Landfill", "material": "Mixed fill",
            "ticket_number": "55501", "tons": "12.4", "cy": ""})
        ws  = openpyxl.load_workbook(dump)["14"]
        hdr = {c.value: c.column for c in ws[5]}
        tons_fill = ws.cell(6, hdr["Tons"]).fill.fgColor.rgb
        note = ws.cell(6, hdr["Helper Notes"]).value
        ok = ws.cell(6, hdr["Ticket #"]).value == "55501" and tons_fill.endswith("FFFF00") \
            and "UNIT CHECK" in (note or "")
        failures += not ok
        print(f"{'PASS' if ok else 'FAIL'}  dump tab 14 row 6: Ticket # "
              f"{ws.cell(6, hdr['Ticket #']).value}, Tons {ws.cell(6, hdr['Tons']).value} "
              f"(fill {tons_fill}), Helper Notes {note!r}")

        rock = shutil.copy(os.path.join(tracker_dir, "ticket_tracker_Permanent Material.xlsx"), tmp)
        r = tp._nf_route(tp._parse_any_qr("01-2610-1111-111111"), lambda k: rock)
        tp.write_to_excel_new_format(rock, "rock", r["toc_row"], {
            "date": "11/10/2026", "company": "Knife River", "material": "4-8in gabion rock",
            "ticket_number": "77001", "tons": "22.1",
            "bid_item_guess": "Gabion rock - likely bid item 200 or 210"}, master_doc="M-00999")
        ws  = openpyxl.load_workbook(rock)["Other Work"]
        hdr = {c.value: c.column for c in ws[5]}
        note = ws.cell(6, hdr["Helper Notes"]).value
        ok = (note == "Gabion rock - likely bid item 200 or 210"
              and ws.cell(6, hdr["Bid Item"]).value is None
              and ws.cell(6, hdr["Master Doc #"]).value == "M-00999")
        failures += not ok
        print(f"{'PASS' if ok else 'FAIL'}  rock Other Work row 6: Helper Notes {note!r}, "
              f"Bid Item left blank, Master Doc # {ws.cell(6, hdr['Master Doc #']).value}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print(f"\n{'ALL PASSED' if not failures else f'{failures} FAILURE(S)'}")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
