"""
setup_2610_trackers.py
======================
Builds the three job 2610 (Beaverton Creek Trail) trackers for the
new-format (type-prefixed) QR codes:

    ticket_tracker_Permanent Material.xlsx  01-2610-…  tabs 7, 8, 9, 10, 13, 16, Other Work
    ticket_tracker_Dump Out.xlsx            02-2610-…  tabs 1, 3, 5, 11, 14, 17, Other Work
    ticket_tracker_Trucking.xlsx            03-2610-…  tabs 2, 4, 6, 12, 15, 18, Other Work

(file names come from NEW_QR_TRACKER_NAMES in ticket_processor.py)

Usage:
    python setup_2610_trackers.py OUT_DIR            # build locally only
    python setup_2610_trackers.py OUT_DIR --upload   # build + replace on SharePoint

--upload first downloads the current 2610 trackers and STOPS if any data tab
has data in it.  It then uploads the three trackers, removes superseded
tracker files (ticket_tracker_2610.xlsx / _rock / _dump / _trucking), and
creates the Dump/Trucking mail folders.  ticket_tracker_2610_concrete.xlsx is left alone.
"""

import os
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

JOB_NUMBER = "2610"
JOB_NAME   = "Beaverton Creek Trail"

ORANGE = PatternFill(start_color="FFFFC000", end_color="FFFFC000", fill_type="solid")  # AI filled
GREY   = PatternFill(start_color="FFD3D3D3", end_color="FFD3D3D3", fill_type="solid")  # human
BOLD   = Font(bold=True, color="FF000000")
TITLE  = Font(bold=True, size=14, color="FF000000")
THIN   = Side(style="thin")
BOX    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP   = Alignment(wrap_text=True, vertical="top")

BID_ITEMS = {
    "350": "Removal of Pipes", "480": "12in Culvert Pipe", "490": "6in Storm Sewer Pipe",
    "500": "12in Storm Sewer Pipe", "510": "24in Storm Sewer Pipe",
    "530": "Concrete Storm Sewer Manholes", "540": "Concrete Manholes with Inlet",
    "550": "Concrete Manholes Large Precast", "560": "Concrete Manholes Water Quality Large",
    "570": "Concrete Inlets Type CG-48 (Mod.)", "580": "Concrete Inlets Type D",
    "590": "Concrete Inlets Type G-2",
    "1280": "Alternate Porous Design from Sta 10+58.7",
    "1290": "Alternate Porous Design from Sta 20+00.5",
    "1300": "Treatment Trench Sta 28+41 to 35+02",
    "1310": "Treatment Trench Sta 35+02 to 40+45",
    "1320": "Treatment Trench Sta 47+81.6 to 55+18.7",
    "1330": "Treatment Trench Sta 69+48.4 to 77+88.3",
    "670": "Prefab Modular Wall (Wall 2)", "680": "Prefab Modular Wall (Wall 3)",
    # Paid by Ticket (pay-quantity) items
    "300": "Contaminated Soil Disposal", "380": "Clearing & Grubbing",
    "390": "General Excavation", "400": "Stone Embankment",
    "440": "Loose Riprap, Class 50", "700": "Aggregate Base",
}

STORM  = "350;480;490;500;510;530;540;550;560"
FRENCH = "1280;1290;1300;1310;1320;1330"
WALLS  = "670;680"

# kind → class, columns, main rows (tab, code name, QR, account, unit, description, bid items),
#        Other Work (QR, accounts, bid item text)
TRACKERS = {
    "rock": {
        "title": "Rock In Ticket Tracking Form",
        "class": "2",
        "columns": ["Ticket Date", "Logged Date", "Supplier/Facility", "Material",
                    "Ticket #", "Tons", "CY", "Helper Notes", "Human Notes"],
        "rows": [
            ("7",  "BI0400-AGG", "01-2610-0400-999999",    "0400",           "tons",
             "Aggregate by the ton", "400"),
            ("8",  "BI0440-AGG", "01-2610-0440-999999",    "0440",           "tons",
             "Aggregate by the ton", "440"),
            ("9",  "BI0700-AGG", "01-2610-0700-999999",    "0700",           "tons",
             "Aggregate by the ton", "700"),
            ("10", "STORM-AGG",  "01-2610-8990-999999-01", "8990-999999-01", "tons",
             "3/4in crushed gravel / aggregate", STORM + ";570"),
            ("13", "FRENCH-AGG", "01-2610-8990-999999-02", "8990-999999-02", "tons",
             "Aggregate by the ton", FRENCH),
            ("16", "WALLS-AGG",  "01-2610-8990-999999-03", "8990-999999-03", "tons",
             "Aggregate by the ton", WALLS),
        ],
        "other": ("01-2610-1111-111111", "0200; 0210; 1020",
                  "200 Construction Entrance Type 1 (acct 0200, gabion rock 4-8in, tons); "
                  "210 Construction Entrance Type 3 (acct 0210, gabion rock 4-8in, tons); "
                  "1020 Bioretention Pond SW Terman Rd (acct 1020, 1in minus drain gravel, CY or tons)"),
    },
    "dump": {
        "title": "Dump Out Ticket Tracking Form",
        "class": "2",
        "columns": ["Ticket Date", "Logged Date", "Dump Site", "Material",
                    "Ticket #", "Tons", "CY", "Helper Notes", "Human Notes"],
        "rows": [
            ("1",  "BI0300-DUMP", "02-2610-0300-999999",   "0300",           "tons",
             "Contaminated soil disposal", "300"),
            ("3",  "BI0380-DUMP", "02-2610-0380-999999",   "0380",           "CY",
             "Dump fees", "380"),
            ("5",  "BI0390-DUMP", "02-2610-0390-999999",   "0390",           "CY",
             "Dump fees", "390"),
            ("11", "STORM-DUMP",  "02-2610-8990-999999-01", "8990-999999-01", "CY",
             "Dump fees", STORM),
            ("14", "FRENCH-DUMP", "02-2610-8990-999999-02", "8990-999999-02", "CY",
             "Dump fees", FRENCH),
            ("17", "WALLS-DUMP",  "02-2610-8990-999999-03", "8990-999999-03", "CY",
             "Material disposal", WALLS),
        ],
        "other": ("02-2610-1111-111111", "0720; 1020",
                  "720 6in Asphalt Pavement Repair (acct 0720, dump fees, CY); "
                  "1020 Bioretention Pond SW Terman Rd (acct 1020, dump fees, CY)"),
    },
    "trucking": {
        "title": "Trucking Ticket Tracking Form",
        "class": "7",
        "columns": ["Ticket Date", "Logged Date", "Trucking Company", "Truck #",
                    "Ticket #", "Loads", "CY", "Tons", "Hours", "Helper Notes", "Human Notes"],
        "rows": [
            ("2",  "BI0300-HAUL", "03-2610-0300-999999",   "0300",           "tons",
             "Trucking sub", "300"),
            ("4",  "BI0380-HAUL", "03-2610-0380-999999",   "0380",           "CY",
             "Trucking sub", "380"),
            ("6",  "BI0390-HAUL", "03-2610-0390-999999",   "0390",           "CY",
             "Trucking sub", "390"),
            ("12", "STORM-HAUL",  "03-2610-8990-999999-01", "8990-999999-01", "CY",
             "Trucking sub / dump truck hauling", STORM + ";570;580;590"),
            ("15", "FRENCH-HAUL", "03-2610-8990-999999-02", "8990-999999-02", "CY",
             "Trucking sub", FRENCH),
            ("18", "WALLS-HAUL",  "03-2610-8990-999999-03", "8990-999999-03", "CY",
             "Trucking sub", WALLS),
        ],
        "other": ("03-2610-1111-111111", "0360; 0720; 1020",
                  "360 Removal of Structures & Obstructions (acct 0360, trucking sub); "
                  "720 6in Asphalt Pavement Repair (acct 0720, trucking sub, CY); "
                  "1020 Bioretention Pond SW Terman Rd (acct 1020, trucking sub, CY)"),
    },
}

HUMAN_COLUMNS = {"Human Notes", "Bid Item"}
TOC_HEADERS   = ["Tab/Sticker #", "Full QR Code", "Code Name", "Description",
                 "Account", "Class", "Unit", "Bid Items Covered"]


def tracker_name(kind: str) -> str:
    """SharePoint file name for a tracker — single source: ticket_processor."""
    from ticket_processor import _nf_tracker_filename
    return _nf_tracker_filename(JOB_NUMBER, kind)


def bid_item_text(codes: str) -> str:
    return "; ".join(f"{c} {BID_ITEMS[c]}" for c in codes.split(";"))


def label(ws, ref, text):
    c = ws[ref]
    c.value, c.font, c.fill, c.border = text, BOLD, GREY, BOX


def value(ws, ref, text):
    c = ws[ref]
    c.value, c.font, c.alignment = text, BOLD, WRAP


def build_tab(wb, cfg, tab, qr, account, unit, bid_items, other_work):
    ws = wb.create_sheet(tab)
    cols = list(cfg["columns"])
    if other_work:
        cols.insert(cols.index("Human Notes"), "Bid Item")
    last = get_column_letter(len(cols))

    ws["A1"] = cfg["title"] + (" — Other Work" if other_work else f" — Tab {tab}")
    ws["A1"].font = TITLE
    ws.row_dimensions[1].height = 20

    label(ws, "A2", "Job Name:");  value(ws, "B2", JOB_NAME)
    label(ws, "E2", "Job Number:"); value(ws, "F2", int(JOB_NUMBER))
    ws.merge_cells("B2:D2")

    label(ws, "A3", "QR Code:");  value(ws, "B3", qr)
    label(ws, "D3", "Account:");  value(ws, "E3", account)
    label(ws, "F3", "Class:");    value(ws, "G3", cfg["class"])
    ws.merge_cells("B3:C3")
    if not other_work:
        label(ws, "H3", "Unit:"); value(ws, "I3", unit)

    label(ws, "A4", "Bid Items Covered:"); value(ws, "B4", bid_items)
    ws.merge_cells(f"B4:{last}4")
    ws.row_dimensions[4].height = 45 if other_work else 32

    for i, name in enumerate(cols, start=1):
        c = ws.cell(row=5, column=i, value=name)
        c.font, c.border = BOLD, BOX
        c.fill = GREY if name in HUMAN_COLUMNS else ORANGE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[5].height = 18.75

    widths = {"Ticket Date": 13, "Logged Date": 13, "Material": 26, "Ticket #": 12,
              "Helper Notes": 40, "Human Notes": 30, "Bid Item": 10}
    for i, name in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(
            name, 22 if name in ("Supplier/Facility", "Dump Site", "Trucking Company") else 10)
    ws.freeze_panes = "A6"


def build_tracker(kind: str, path: str) -> None:
    cfg = TRACKERS[kind]
    wb  = openpyxl.Workbook()
    toc = wb.active
    toc.title = "TOC"
    for i, h in enumerate(TOC_HEADERS, start=1):
        c = toc.cell(row=1, column=i, value=h)
        c.font, c.fill, c.border = BOLD, GREY, BOX
    # Tabs in numeric order, Other Work last (TOC rows in the same order)
    for tab, code, qr, acct, unit, desc, bids in sorted(cfg["rows"], key=lambda r: int(r[0])):
        bids_txt = bid_item_text(bids)
        toc.append([tab, qr, code, desc, acct, cfg["class"], unit, bids_txt])
        build_tab(wb, cfg, tab, qr, acct, unit, bids_txt, other_work=False)
    qr, accts, bids_txt = cfg["other"]
    toc.append(["Other Work", qr, "OTHER-WORK", "Other Work", accts, cfg["class"],
                "per bid item", bids_txt])
    build_tab(wb, cfg, "Other Work", qr, accts, "", bids_txt, other_work=True)

    for col, w in zip("ABCDEFGH", (11, 26, 14, 34, 18, 7, 12, 90)):
        toc.column_dimensions[col].width = w
    for row in toc.iter_rows(min_row=2):
        for c in row:
            c.alignment = WRAP
    toc.freeze_panes = "A2"
    wb.save(path)


def _has_data(path: str, data_start_row: int) -> list[str]:
    """Return 'tab!cell' for every non-empty cell in a data area (TOC excluded)."""
    found = []
    wb = openpyxl.load_workbook(path, read_only=True)
    for ws in wb.worksheets:
        if ws.title == "TOC":
            continue
        for row in ws.iter_rows(min_row=data_start_row):
            for c in row:
                if c.value not in (None, ""):
                    found.append(f"{ws.title}!{c.coordinate}={c.value!r}")
    wb.close()
    return found


def upload(out_dir: str) -> None:
    import requests
    import ticket_processor as tp

    client   = tp.GraphClient()
    site_id  = tp._sharepoint_site_id(client)
    drive_id = tp._sharepoint_drive_id(client, site_id)
    job      = tp._find_job_folder_by_prefix(client, JOB_NUMBER)
    if not job:
        sys.exit(f"No SharePoint folder starting with {JOB_NUMBER} — aborting.")
    base = f"{tp.GRAPH_BASE}/sites/{site_id}/drives/{drive_id}/root:/{tp.SHAREPOINT_FOLDER}/{job}"
    existing = {i["name"] for i in client.get(f"{base}:/children").json()["value"]}

    # Safety check: refuse to replace any tracker whose data tabs hold data.
    check_dir = os.path.join(out_dir, "current_sharepoint")
    os.makedirs(check_dir, exist_ok=True)
    current_names = {tp._nf_tracker_filename(JOB_NUMBER, k) for k in TRACKERS}
    superseded    = {"ticket_tracker_2610.xlsx", "ticket_tracker_2610_rock.xlsx",
                     "ticket_tracker_2610_dump.xlsx", "ticket_tracker_2610_trucking.xlsx"}
    old_layout = {n: 6 for n in current_names | superseded}
    old_layout["ticket_tracker_2610.xlsx"] = 9               # old 2602-style layout
    for name, start in old_layout.items():
        if name not in existing:
            continue
        local = os.path.join(check_dir, name)
        if not tp._download_excel_file(client, name, local, job_folder=job):
            sys.exit(f"Could not download {name} to check it — aborting.")
        # A file already in the new layout has headers on row 5, data from row 6.
        wb = openpyxl.load_workbook(local, read_only=True)
        new_layout = any(ws.title in ("Other Work",) for ws in wb.worksheets)
        wb.close()
        data = _has_data(local, 6 if new_layout else start)
        if data:
            sys.exit(f"STOP: {name} has data — not replacing it:\n  " + "\n  ".join(data[:20]))
        print(f"checked {name}: no data")

    for kind in TRACKERS:
        name = tracker_name(kind)
        tp._upload_excel_file(client, name, os.path.join(out_dir, name), job_folder=job)
        print(f"uploaded {name}")

    for old in sorted(superseded & existing):
        client._ensure_fresh_token()
        r = requests.delete(f"{base}/{old}", headers=client._json_headers())
        r.raise_for_status()
        print(f"removed superseded {old}")

    for t in ("dump", "trucking"):
        tp.get_or_create_processed_folder(client, t)
        tp.get_or_create_review_folder(client, t)
        print(f"mail folders ready: {tp._processed_folder_name(t)!r}, {tp._review_folder_name(t)!r}")

    after = sorted(i["name"] for i in client.get(f"{base}:/children").json()["value"])
    print("SharePoint folder now contains:", after)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    out = sys.argv[1]
    os.makedirs(out, exist_ok=True)
    for kind in TRACKERS:
        p = os.path.join(out, tracker_name(kind))
        build_tracker(kind, p)
        print("built", p)
    if "--upload" in sys.argv:
        upload(out)
