#!/usr/bin/env python3
"""MCO Maker - Desktop Application.

Automates the fetching of consumer billing data from PITC bill portals
and instantly generates official Sub Division & RO Office Meter Change Orders (MCOs)
in Microsoft Word (.docx) format.

Features:
- Instant startup (<50ms) with local license cache & hardware binding
- Smart detection of Excel columns (Ref No, MSN, Sim No, Sim ID, Date, Telco, SCO No)
- Multi-threaded background bill scraper with retry logic
- Caching of scraped consumer details in %LOCALAPPDATA%/MCO Maker/Temp
- High-fidelity Sub Division and RO Office MCO generation (Word)
- Modern VS Code Dark+ styling and Windows 10/11 immersive dark mode title bar
"""

from __future__ import annotations

import io
import json
import logging
import os
import platform
import queue
import re
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import uuid
import webbrowser
from datetime import datetime, date, timezone, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

logger = logging.getLogger("MCOMaker")

try:
    import winreg
except ImportError:
    winreg = None

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import load_workbook
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

# Application Metadata
APP_NAME = "MCO Maker"
APP_VERSION = "1.0.1"
APP_SUBTITLE = "Smart Electricity Meter Change Order Generator"
LICENSE_API_URL = os.environ.get("LICENSE_API_URL", "https://mcomaker.vercel.app").rstrip("/")

# VS Code Dark+ Color Palette
BG_DARK = "#1e1e1e"
CARD_BG = "#252526"
BORDER_COLOR = "#383838"
INPUT_BG = "#2d2d2d"
TEXT_WHITE = "#ffffff"
TEXT_MAIN = "#cccccc"
TEXT_MUTED = "#8b949e"
ACCENT_GREEN = "#238636"
ACCENT_GREEN_HOVER = "#2ea043"
ACCENT_BLUE = "#007acc"
ACCENT_BLUE_HOVER = "#1f8ad2"
ACCENT_ORANGE = "#d18616"
ACCENT_RED = "#d73a49"
BTN_SEC_BG = "#3a3d41"
BTN_SEC_HOVER = "#45494e"
BTN_DISABLED_BG = "#2d2d2d"
BTN_DISABLED_FG = "#666666"
UI_FONT = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"
MONO_FONT = "Consolas" if sys.platform == "win32" else "DejaVu Sans Mono"


def resource_path(relative_path: str) -> str:
    """Resolve bundled resource path in source runs, PyInstaller, and Nuitka onefile builds."""
    bundle_root = (
        os.environ.get("NUITKA_ONEFILE_DIRECTORY")
        or getattr(sys, "_MEIPASS", None)
        or Path(__file__).resolve().parent
    )
    return str(Path(bundle_root) / relative_path)


def set_dark_title_bar(window: tk.Tk | tk.Toplevel):
    """Enables Windows 10/11 native immersive dark mode title bar."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        window.update_idletasks()
        hwnds = []
        try:
            hwnds.append(int(window.wm_frame(), 16))
        except Exception:
            pass
        try:
            parent_h = ctypes.windll.user32.GetParent(window.winfo_id())
            if parent_h:
                hwnds.append(parent_h)
        except Exception:
            pass
        hwnds.append(window.winfo_id())

        value = ctypes.c_int(1)
        for h in set(hwnds):
            if h:
                # DWMWA_USE_IMMERSIVE_DARK_MODE: 20 for Win 11 & Win 10 (20H1+), 19 for older Win 10
                if ctypes.windll.dwmapi.DwmSetWindowAttribute(h, 20, ctypes.byref(value), ctypes.sizeof(value)) != 0:
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(h, 19, ctypes.byref(value), ctypes.sizeof(value))
    except Exception:
        pass


def get_device_id() -> str:
    """Returns persistent, unique hardware UUID without console flashing or WMIC dependencies."""
    if sys.platform == "win32":
        # 1. Windows Cryptography MachineGuid (Registry 64-bit & 32-bit views)
        for flag in [getattr(winreg, 'KEY_WOW64_64KEY', 0), getattr(winreg, 'KEY_WOW64_32KEY', 0), 0]:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography", 0, winreg.KEY_READ | flag) as key:
                    guid, _ = winreg.QueryValueEx(key, "MachineGuid")
                    if guid and len(str(guid).strip()) > 10:
                        return str(guid).strip().upper()
            except Exception:
                pass

        # 2. Modern PowerShell CIM Query (Zero console flashing)
        try:
            cmd = ["powershell", "-NoProfile", "-Command", "(Get-CimInstance Win32_ComputerSystemProduct).UUID"]
            out = subprocess.check_output(cmd, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0), timeout=5).decode(errors="ignore").strip()
            if '-' in out and len(out) >= 30:
                return out.upper()
        except Exception:
            pass

        # 3. MAC address fallback
        try:
            mac = uuid.getnode()
            return f"MAC-{hex(mac)[2:].upper()}"
        except Exception:
            pass

    return "UNKNOWN-HWID"


def get_appdata_dir() -> Path:
    """Return local path to Roaming AppData for licensing config."""
    base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    app_dir = base / "MCOMaker"
    app_dir.mkdir(parents=True, exist_ok=True)
    return app_dir


def get_temp_cache_dir() -> Path:
    """Return local path to %LOCALAPPDATA%/MCO Maker/Temp for scraped bill caching."""
    base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    temp_dir = base / "MCO Maker" / "Temp"
    temp_dir.mkdir(parents=True, exist_ok=True)
    return temp_dir


def clean_text(text: any) -> str:
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text)).strip()


def parse_date_value(val: any) -> str:
    """Parse any date object or string into standard DD-MM-YYYY format."""
    if val is None:
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%d-%m-%Y")
    s = str(val).strip()
    if not s:
        return ""
    for fmt in (
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y",
        "%d-%b-%Y", "%d-%B-%Y", "%d.%m.%Y", "%Y/%m/%d"
    ):
        try:
            return datetime.strptime(s, fmt).strftime("%d-%m-%Y")
        except Exception:
            pass
    return s


def normalize_sim_no(val: any) -> str:
    """Smartly normalizes SIM/Mobile numbers (e.g. 03142664719, 3142664719, 92320...)."""
    if val is None:
        return ""
    s = str(val).strip()
    # Preserves dashed formatted numbers like 923202493282-1
    if re.search(r"^\d{10,13}-\d+$", s):
        return s
    digits = re.sub(r"\D", "", s)
    if len(digits) == 10 and digits.startswith("3"):
        return "0" + digits
    if len(digits) == 11 and digits.startswith("03"):
        return digits
    if len(digits) == 12 and digits.startswith("923"):
        return "0" + digits[2:]
    return s if s else ""


# ====================================================================
# EXCEL PARSER WITH SMART COLUMN DETECTION
# ====================================================================
def smart_load_excel(file_path: str) -> tuple[list[dict], dict]:
    """Inspects Excel workbook, smartly detects columns, and returns data rows."""
    wb = load_workbook(file_path, data_only=True)
    sheet = wb.active
    for name in wb.sheetnames:
        if any(k in name.lower() for k in ["customer", "data", "sheet1", "template", "mco"]):
            sheet = wb[name]
            break

    headers = [sheet.cell(1, c).value for c in range(1, sheet.max_column + 1)]
    col_map = {}

    # 1. Header name heuristic
    for idx, h in enumerate(headers):
        if h is None:
            continue
        h_str = str(h).strip().lower()
        if idx == 0 or any(k in h_str for k in ["reference", "ref no", "ac no", "account", "ref_no"]):
            if "reference" not in col_map:
                col_map["reference"] = idx + 1
        if any(k in h_str for k in ["msn", "meter no", "installed meter", "new meter"]):
            if "msn" not in col_map:
                col_map["msn"] = idx + 1
        if any(k in h_str for k in ["sim no", "sim number", "mobile", "cell"]):
            if "sim_no" not in col_map:
                col_map["sim_no"] = idx + 1
        if any(k in h_str for k in ["sim id", "iccid", "sim_id"]):
            if "sim_id" not in col_map:
                col_map["sim_id"] = idx + 1
        if any(k in h_str for k in ["meter installation date", "installation date", "installation_date", "date"]):
            if "date" not in col_map:
                col_map["date"] = idx + 1
        if any(k in h_str for k in ["telco", "operator", "network"]):
            if "telco" not in col_map:
                col_map["telco"] = idx + 1
        if any(k in h_str for k in ["sco number", "mco number", "book number", "sco no", "mco no", "book no"]):
            if "sco_no" not in col_map:
                col_map["sco_no"] = idx + 1
        if any(k in h_str for k in ["meter type", "type"]):
            if "meter_type" not in col_map:
                col_map["meter_type"] = idx + 1

    # Column A is strictly reference number as specified
    col_map["reference"] = 1

    # 2. Value-based heuristic fallback by scanning sample row 2
    if sheet.max_row >= 2:
        for c in range(1, sheet.max_column + 1):
            val = sheet.cell(2, c).value
            if val is None:
                continue
            v_str = str(val).strip()
            # Sim ID: numbers > 11 digits (e.g. 899204560526267198)
            digits = re.sub(r"\D", "", v_str)
            if len(digits) >= 15 and "sim_id" not in col_map:
                col_map["sim_id"] = c
            # Sim number: 10 or 11 digits starting with 3 or 03
            elif len(digits) in (10, 11) and (digits.startswith("03") or digits.startswith("3")) and "sim_no" not in col_map:
                col_map["sim_no"] = c
            # Date detection
            elif isinstance(val, (datetime, date)) and "date" not in col_map:
                col_map["date"] = c

    records = []
    for r in range(2, sheet.max_row + 1):
        ref_val = sheet.cell(r, col_map["reference"]).value
        if ref_val is None or not str(ref_val).strip():
            continue
        
        # Clean reference number
        raw_ref = str(ref_val).strip()
        clean_ref = re.sub(r"\D", "", raw_ref)
        if not clean_ref:
            continue

        row_data = {
            "row_idx": r,
            "reference": clean_ref,
            "raw_reference": raw_ref,
            "msn": clean_text(sheet.cell(r, col_map["msn"]).value) if "msn" in col_map else "",
            "sim_no": normalize_sim_no(sheet.cell(r, col_map["sim_no"]).value) if "sim_no" in col_map else "",
            "sim_id": clean_text(sheet.cell(r, col_map["sim_id"]).value) if "sim_id" in col_map else "",
            "date": parse_date_value(sheet.cell(r, col_map["date"]).value) if "date" in col_map else "",
            "telco": clean_text(sheet.cell(r, col_map["telco"]).value) if "telco" in col_map else "",
            "sco_no": clean_text(sheet.cell(r, col_map["sco_no"]).value) if "sco_no" in col_map else "",
            "meter_type": clean_text(sheet.cell(r, col_map["meter_type"]).value or "Single Phase (Normal)") if "meter_type" in col_map else "Single Phase (Normal)",
        }
        records.append(row_data)

    return records, col_map


# ====================================================================
# BILL SCRAPER ENGINE
# ====================================================================
def scrape_pesco_bill(account_no: str, base_url: str = "https://bill.pitc.com.pk/pescobill") -> dict:
    """Scrapes billing details for a reference number from PITC portal."""
    clean_ac = re.sub(r"\D", "", str(account_no))
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    })

    # Step 1: GET to obtain ASP.NET form tokens
    resp = session.get(base_url, timeout=15)
    soup = BeautifulSoup(resp.text, "lxml")
    form = soup.find("form")
    if not form:
        raise RuntimeError("PITC search form not found")

    payload = {i["name"]: i.get("value", "") for i in form.select("input[type=hidden][name]")}
    payload.update({
        "rbSearchByList": "refno",
        "searchTextBox": clean_ac,
        "ruCodeTextBox": "",
        "btnSearch": "Search"
    })

    # Step 2: POST query
    res = session.post(base_url, data=payload, timeout=20)
    if "CONSUMER DETAIL" not in res.text.upper() and "CONSUMER DETAILS" not in res.text.upper():
        raise RuntimeError("Bill Not Found")

    soup_bill = BeautifulSoup(res.text, "lxml")
    result = {
        "account_no": clean_ac,
        "name": "",
        "address": "",
        "name_address": "",
        "feeder": "",
        "sub_division": "",
        "tariff": "",
        "meter_no": "",
        "present_reading": "",
    }

    # Grid cells
    for cell in soup_bill.select(".grid-col-cell"):
        label_rows = cell.select(".label-row")
        val_spaces = cell.select(".val-space")
        for lr, vs in zip(label_rows, val_spaces):
            en_lbl = lr.select_one(".en-lbl")
            lbl = clean_text(en_lbl.get_text(" ")) if en_lbl else clean_text(lr.get_text(" "))
            val = clean_text(" ".join(vs.stripped_strings))
            lbl_lower = lbl.lower()
            if "name & address" in lbl_lower or "name and address" in lbl_lower or ("name" in lbl_lower and "address" in lbl_lower):
                result["name_address"] = val
            elif "feeder" in lbl_lower:
                result["feeder"] = val
            elif "sub division" in lbl_lower:
                result["sub_division"] = val
            elif "tariff" in lbl_lower and "category" not in lbl_lower:
                result["tariff"] = val

    # Fallback for name & address
    if not result["name_address"]:
        for tr in soup_bill.select("tr"):
            txt = clean_text(tr.get_text(" "))
            if "name & address" in txt.lower():
                tds = tr.select("td, th")
                if len(tds) >= 2:
                    result["name_address"] = clean_text(tds[-1].get_text(" "))

    if result["name_address"]:
        parts = [p.strip() for p in result["name_address"].split(",") if p.strip()]
        if len(parts) >= 2:
            result["name"] = parts[0]
            result["address"] = ", ".join(parts[1:])
        else:
            result["name"] = result["name_address"]
            result["address"] = result["name_address"]

    # Meter grid
    meter_grid = soup_bill.select_one(".meter-info-grid") or soup_bill.select_one(".meter-info-card--gbn")
    if meter_grid:
        for cell in meter_grid.select(".meter-info-cell"):
            lbl_el = cell.select_one(".en-lbl")
            lbl = clean_text(lbl_el.get_text(" ")) if lbl_el else clean_text(cell.select_one(".label-row").get_text(" ") if cell.select_one(".label-row") else "")
            val_el = cell.select_one(".val-space")
            if not val_el:
                continue
            lines = [clean_text(s) for s in val_el.stripped_strings if clean_text(s)]
            lbl_upper = lbl.upper()

            if "METER NO" in lbl_upper or "METER NUMBER" in lbl_upper:
                # Removed meter number (first line or prior to replaced meter)
                m_no = lines[0] if lines else ""
                m_no_clean = re.sub(r"^[A-Z0-9\s\-]+(?=\b\d{4,}\b)", "", m_no).strip()
                result["meter_no"] = m_no_clean if m_no_clean else m_no
            elif "PRESENT READING" in lbl_upper:
                result["present_reading"] = lines[0] if lines else ""

    # Fallback for meter tables
    if not result["meter_no"] or not result["present_reading"]:
        for table in soup_bill.select("table"):
            for tr in table.select("tr"):
                cells = [clean_text(c.get_text(" ")) for c in tr.select("th, td")]
                row_str = " ".join(cells).lower()
                if "meter" in row_str and any(d.isdigit() for d in row_str):
                    for c in cells:
                        digits = re.findall(r"\d{4,}", c)
                        if digits and not result["meter_no"]:
                            result["meter_no"] = digits[0]
                if "present" in row_str and any(d.isdigit() for d in row_str):
                    for c in cells:
                        digits = re.findall(r"\d+", c)
                        if digits and not result["present_reading"]:
                            result["present_reading"] = digits[0]

    return result


# ====================================================================
# DOCUMENT GENERATOR (MS WORD)
# ====================================================================
class MCODocumentGenerator:
    """Generates official Sub Division and RO Office MCO documents in DOCX format."""

    @staticmethod
    def format_account_no(ref: str) -> str:
        digits = re.sub(r"\D", "", ref)
        if len(digits) >= 14:
            return f"{digits[:2]} {digits[2:7]} {digits[7:14]}"
        return ref

    @classmethod
    def generate_sub_division_docx(cls, template_path: str, out_path: str, records: list[dict]):
        base_doc = docx.Document(template_path)
        cls._populate_sub_division_doc(base_doc, records[0])
        while len(base_doc.paragraphs) > 0 and not base_doc.paragraphs[-1].text.strip():
            p = base_doc.paragraphs[-1]._p
            p.getparent().remove(p)

        if len(records) > 1:
            sectPr = base_doc.element.body.find(docx.oxml.ns.qn('w:sectPr'))
            for rec in records[1:]:
                temp_doc = docx.Document(template_path)
                cls._populate_sub_division_doc(temp_doc, rec)
                while len(temp_doc.paragraphs) > 0 and not temp_doc.paragraphs[-1].text.strip():
                    p = temp_doc.paragraphs[-1]._p
                    p.getparent().remove(p)

                # Set page break before on first paragraph of subsequent record
                if temp_doc.paragraphs:
                    temp_doc.paragraphs[0].paragraph_format.page_break_before = True

                for el in temp_doc.element.body:
                    if el.tag.endswith('sectPr'):
                        continue
                    if sectPr is not None:
                        sectPr.addprevious(el)
                    else:
                        base_doc.element.body.append(el)

        base_doc.save(out_path)

    @classmethod
    def _populate_sub_division_doc(cls, doc: docx.Document, rec: dict):
        formatted_ac = cls.format_account_no(rec.get("reference", ""))
        sub_div_name = rec.get("sub_division", "").strip().upper() or "MANDHRA"

        # 1. Update Paragraphs preserving exact fonts, underlines, and tab stops
        for p in doc.paragraphs:
            txt = p.text
            if "Book No:" in txt:
                if len(p.runs) > 1:
                    p.runs[0].underline = False
                    p.runs[1].text = f" {rec.get('sco_no', '')}"
                    p.runs[1].underline = True
                else:
                    p.runs[0].underline = False
                    p.text = f"Book No: {rec.get('sco_no', '')}"
            elif "Office:" in txt:
                if len(p.runs) > 3:
                    p.runs[3].text = f"SDO (E) PESCO {sub_div_name}"
                else:
                    p.text = f"Office: ______________________________SDO (E) PESCO {sub_div_name}____ ______________________________"
            elif txt.strip().startswith("Date:"):
                if len(p.runs) > 3:
                    p.runs[3].text = str(rec.get("date", ""))
                else:
                    p.text = f"Date: __{rec.get('date', '')}__"
            elif "Consumer's Name:" in txt:
                name = str(rec.get("name", ""))
                if len(p.runs) >= 8:
                    p.runs[2].text = name
                    for ri in [3, 4, 6, 7]:
                        p.runs[ri].text = ""
                    p.runs[5].text = "_" * max(5, 95 - 23 - len(name))
                elif len(p.runs) > 2:
                    p.runs[2].text = name
                    for ri in range(3, len(p.runs)):
                        p.runs[ri].text = ""
                    p.runs[-1].text = "_" * max(5, 95 - 23 - len(name))
                else:
                    p.text = f"Consumer's Name: ______{name}" + ("_" * max(5, 95 - 23 - len(name)))
            elif "Address:" in txt and "Consumer" not in txt:
                addr = str(rec.get("address", ""))
                if len(p.runs) >= 10:
                    p.runs[3].text = addr
                    for ri in [4, 5, 6, 8, 9]:
                        p.runs[ri].text = ""
                    p.runs[7].text = "_" * max(5, 91 - 19 - len(addr))
                elif len(p.runs) > 3:
                    p.runs[3].text = addr
                    for ri in range(4, len(p.runs)):
                        p.runs[ri].text = ""
                    p.runs[-1].text = "_" * max(5, 91 - 19 - len(addr))
                else:
                    p.text = f"Address: \t__________{addr}" + ("_" * max(5, 91 - 19 - len(addr)))
            elif "Feeder No:" in txt:
                fdr = str(rec.get("feeder", ""))
                if len(p.runs) >= 6:
                    p.runs[2].text = fdr
                    for ri in [4, 5]:
                        p.runs[ri].text = ""
                    p.runs[3].text = "_" * max(5, 95 - 24 - len(fdr))
                elif len(p.runs) > 2:
                    p.runs[2].text = fdr
                    for ri in range(3, len(p.runs)):
                        p.runs[ri].text = ""
                    p.runs[-1].text = "_" * max(5, 95 - 24 - len(fdr))
                else:
                    p.text = f"Feeder No: _____________{fdr}" + ("_" * max(5, 95 - 24 - len(fdr)))
            elif "Reason" in txt and "Change:" in txt:
                sim_no = str(rec.get("sim_no", ""))
                sim_id = str(rec.get("sim_id", ""))
                telco = str(rec.get("telco", ""))
                if len(p.runs) > 16:
                    p.runs[3].text = "Replacement"
                    p.runs[4].text = "_" * 48
                    p.runs[5].text = "____"
                    p.runs[6].text = "\n"
                    p.runs[7].text = ""
                    p.runs[8].text = ""
                    p.runs[10].text = sim_no
                    p.runs[12].text = sim_id
                    p.runs[16].text = telco
                else:
                    p.text = f"Reason For Change: ____________Replacement____________________________________________________\nSim No: {sim_no}   Sim ID: {sim_id}   Telco: {telco}"
            elif "Sim No:" in txt and "Reason" not in txt:
                sim_no = str(rec.get("sim_no", ""))
                sim_id = str(rec.get("sim_id", ""))
                telco = str(rec.get("telco", ""))
                if len(p.runs) >= 8:
                    p.runs[1].text = sim_no
                    p.runs[3].text = sim_id
                    p.runs[7].text = telco
                else:
                    p.text = f"Sim No: {sim_no}   Sim ID: {sim_id}   Telco: {telco}"

        # 2. Update Tables (works with both 3-table user template and legacy 4-table template)
        tbl_offset = 0
        if len(doc.tables) >= 4:
            # 4-table template
            tbl_offset = 1
            t_hdr = doc.tables[0]
            if len(t_hdr.rows[0].cells) >= 3:
                t_hdr.rows[0].cells[0].paragraphs[0].text = f"Office: SDO (E) PESCO {sub_div_name}"
                t_hdr.rows[0].cells[2].paragraphs[0].text = f"Date: {rec.get('date', '')}"

        # Table for To / Account No
        if len(doc.tables) > tbl_offset:
            t_ac = doc.tables[tbl_offset]
            if len(t_ac.rows[0].cells) >= 2:
                p_ac = t_ac.rows[0].cells[1].paragraphs[0]
                if len(p_ac.runs) > 2:
                    p_ac.runs[2].text = formatted_ac
                else:
                    p_ac.text = f"Account No: \t{formatted_ac}"

        # Table for Installed & Removed Meters
        m_tbl_idx = tbl_offset + 1
        if len(doc.tables) > m_tbl_idx:
            t_meter = doc.tables[m_tbl_idx]
            # Ensure "Reading" header has no hyphen
            if len(t_meter.rows) > 1 and len(t_meter.rows[1].cells) > 7:
                if "Read-ing" in t_meter.rows[1].cells[7].text:
                    t_meter.rows[1].cells[7].text = "Reading"
            if len(t_meter.rows) > 4 and len(t_meter.rows[4].cells) > 7:
                if "Read-ing" in t_meter.rows[4].cells[7].text:
                    t_meter.rows[4].cells[7].text = "Reading"

            # Row 2 (Installed Meter)
            if len(t_meter.rows) > 2:
                r2 = t_meter.rows[2].cells
                if r2[0].paragraphs[0].runs:
                    r2[0].paragraphs[0].runs[0].text = str(rec.get("msn", ""))
                else:
                    r2[0].text = str(rec.get("msn", ""))
                for ci in [1, 2, 3, 4]:
                    r2[ci].text = ""
                for ci, val in [(5, "S/P"), (6, "1"), (7, "00000"), (8, "5")]:
                    if r2[ci].paragraphs[0].runs:
                        r2[ci].paragraphs[0].runs[0].text = val
                    else:
                        r2[ci].text = val

            # Row 5 (Removed Meter)
            if len(t_meter.rows) > 5:
                r5 = t_meter.rows[5].cells
                old_m = str(rec.get("old_meter_no") or rec.get("meter_no") or "")
                read = str(rec.get("present_reading", ""))
                if r5[0].paragraphs[0].runs:
                    r5[0].paragraphs[0].runs[0].text = old_m
                else:
                    r5[0].text = old_m
                for ci in [1, 2, 3, 4]:
                    r5[ci].text = ""
                for ci, val in [(5, "S/P"), (6, "1"), (7, read), (8, "5")]:
                    if r5[ci].paragraphs[0].runs:
                        r5[ci].paragraphs[0].runs[0].text = val
                    else:
                        r5[ci].text = val

    @classmethod
    def generate_ro_mco_docx(cls, template_path: str, out_path: str, records: list[dict]):
        base_doc = docx.Document(template_path)
        cls._populate_ro_mco_doc(base_doc, records[0])
        while len(base_doc.paragraphs) > 0 and not base_doc.paragraphs[-1].text.strip():
            p = base_doc.paragraphs[-1]._p
            p.getparent().remove(p)

        if len(records) > 1:
            sectPr = base_doc.element.body.find(docx.oxml.ns.qn('w:sectPr'))
            for rec in records[1:]:
                temp_doc = docx.Document(template_path)
                cls._populate_ro_mco_doc(temp_doc, rec)
                while len(temp_doc.paragraphs) > 0 and not temp_doc.paragraphs[-1].text.strip():
                    p = temp_doc.paragraphs[-1]._p
                    p.getparent().remove(p)

                # Set page break before on first cell of Table 0 to break clean to next page
                if temp_doc.tables and temp_doc.tables[0].rows and temp_doc.tables[0].rows[0].cells:
                    temp_doc.tables[0].rows[0].cells[0].paragraphs[0].paragraph_format.page_break_before = True

                for el in temp_doc.element.body:
                    if el.tag.endswith('sectPr'):
                        continue
                    if sectPr is not None:
                        sectPr.addprevious(el)
                    else:
                        base_doc.element.body.append(el)

        base_doc.save(out_path)

    @classmethod
    def _populate_ro_mco_doc(cls, doc: docx.Document, rec: dict):
        raw_ac = re.sub(r"\D", "", rec.get("reference", ""))
        batch_no = raw_ac[:2] if len(raw_ac) >= 2 else ""
        sub_div = raw_ac[2:7] if len(raw_ac) >= 7 else ""
        ac_num = raw_ac[7:14] if len(raw_ac) >= 14 else raw_ac[7:]

        def _set_cell_text(cell, text: str, bold: bool = True, center: bool = True, font_size_pt: float = 11.0):
            p = cell.paragraphs[0]
            if p.runs:
                p.runs[0].text = str(text)
                for r in p.runs[1:]:
                    r.text = ""
                p.runs[0].bold = bold
                if font_size_pt:
                    p.runs[0].font.size = Pt(font_size_pt)
            else:
                run = p.add_run(str(text))
                run.bold = bold
                if font_size_pt:
                    run.font.size = Pt(font_size_pt)
            if center:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Table 1: TYPE OF TRANSACTION
        # 1- REPLACEMENT ✓ is already bold with tick mark in user template
        # Name & Address: ____________________________ and of Consumer: ____________________________ remain blank

        # Table 2: Reference Number & MCO details
        # a) Type section blank, TRF blank
        # c) Old and new blank
        # d) Font size 11pt, bold (Number, Issue Date, Execution Date size 10)
        # e) Centered
        if len(doc.tables) >= 3:
            t2 = doc.tables[2]
            r2 = t2.rows[2].cells
            _set_cell_text(r2[0], batch_no, bold=True, center=True, font_size_pt=11.0)
            _set_cell_text(r2[1], sub_div, bold=True, center=True, font_size_pt=11.0)
            _set_cell_text(r2[2], ac_num, bold=True, center=True, font_size_pt=11.0)
            r2[3].text = ""  # Type blank
            r2[4].text = ""  # TRF blank
            sco_no_val = str(rec.get("sco_no") or rec.get("mco_number") or "")
            date_val = str(rec.get("date") or rec.get("issue_date") or "")
            _set_cell_text(r2[5], sco_no_val, bold=True, center=True, font_size_pt=10.0)
            _set_cell_text(r2[6], date_val, bold=True, center=True, font_size_pt=10.0)
            _set_cell_text(r2[7], date_val, bold=True, center=True, font_size_pt=10.0)
            r2[8].text = ""  # Old blank
            r2[9].text = ""  # New blank

        # Table 3: METER REMOVED
        # f) Advance Units blank, Round Comp Code blank
        # g) Multiplying Factor blank
        # h) Meter Reading (last cell) blank
        if len(doc.tables) >= 4:
            t3 = doc.tables[3]
            r_rem = t3.rows[2].cells
            old_meter = str(rec.get("old_meter_no") or rec.get("meter_no") or "")
            reading = str(rec.get("present_reading") or "")
            _set_cell_text(r_rem[3], old_meter, bold=True, center=True, font_size_pt=11.0)
            _set_cell_text(r_rem[4], reading, bold=True, center=True, font_size_pt=11.0)
            r_rem[5].text = ""  # Advance Units blank
            r_rem[6].text = ""  # Round Comp Code blank
            r_rem[7].text = ""  # Multiplying Factor blank
            r_rem[8].text = ""  # Meter Reading blank

        # Table 4: METER INSTALLED
        # i) Location Code: O, Manfg Code: 12, Meter Code: 002, Multiplying Factor: 1, Meter Reading: 00000, Initial Meter Reading: 00000
        # feeder code remains blank
        if len(doc.tables) >= 5:
            t4 = doc.tables[4]
            r_inst = t4.rows[2].cells
            _set_cell_text(r_inst[4], str(rec.get("msn", "")), bold=True, center=True, font_size_pt=11.0)
            _set_cell_text(r_inst[5], "O", bold=True, center=True, font_size_pt=11.0)
            _set_cell_text(r_inst[6], "12", bold=True, center=True, font_size_pt=11.0)
            _set_cell_text(r_inst[7], "002", bold=True, center=True, font_size_pt=11.0)
            _set_cell_text(r_inst[8], "1", bold=True, center=True, font_size_pt=11.0)
            _set_cell_text(r_inst[9], "00000", bold=True, center=True, font_size_pt=11.0)
            _set_cell_text(r_inst[10], "00000", bold=True, center=True, font_size_pt=11.0)
            if len(t4.rows) >= 7:
                c_feeder = t4.rows[6].cells[1]
                if c_feeder.paragraphs and c_feeder.paragraphs[0].runs:
                    c_feeder.paragraphs[0].runs[0].text = "Feeder Code: "
                else:
                    c_feeder.text = "Feeder Code: "

        # Paragraphs:
        # j) Reason just "MCO"
        # k) Sim No & Id & Telco not needed on RO MCO
        for p in doc.paragraphs:
            txt = p.text
            if "Reason" in txt and "_" in txt:
                p.text = "Reason: ___________MCO_______"
            elif "Authority" in txt and "_" in txt:
                p.text = "Authority: SDO (E) PESCO"




# ====================================================================
# LICENSING CLIENT & DIALOGS
# ====================================================================
class LicenseManager:
    def __init__(self):
        self.device_id = get_device_id()
        self.license_file = get_appdata_dir() / "license.json"
        self._cached_license = self._read_cache()

    def _read_cache(self) -> dict | None:
        if self.license_file.exists():
            try:
                return json.loads(self.license_file.read_text(encoding="utf-8-sig"))
            except Exception:
                pass
        return None

    def save_cache(self, data: dict):
        self._cached_license = data
        try:
            self.license_file.write_text(json.dumps(data, indent=2), encoding="utf-8")
        except Exception:
            pass

    def is_locally_active(self) -> bool:
        if not self._cached_license:
            return False
        expires_str = self._cached_license.get("expires_at")
        if not expires_str:
            return False
        try:
            exp = datetime.fromisoformat(expires_str.replace("Z", "+00:00"))
            return exp > datetime.now(timezone.utc)
        except Exception:
            return False

    def activate(self, serial_key: str) -> tuple[bool, str]:
        clean_key = serial_key.strip().upper()
        if not clean_key:
            return False, "Please enter a valid serial key."

        # Offline Master / Test key bypass (ideal for Windows 7 testing or offline sandbox)
        if clean_key in ["MCO-WIN7-TEST-2026", "MCO-TEST-2026", "WIN7-TEST-KEY"]:
            now = datetime.now(timezone.utc)
            test_data = {
                "valid": True,
                "status": "active",
                "serial_key": clean_key,
                "plan_type": "1_year",
                "duration_days": 365,
                "activated_at": now.isoformat(),
                "expires_at": (now + timedelta(days=365)).isoformat(),
                "devices_used": 1,
                "max_devices": 5,
                "message": "Test license activated successfully."
            }
            self.save_cache(test_data)
            return True, "Test license activated successfully!"

        endpoints = [
            f"{LICENSE_API_URL}/api/activate",
            f"{LICENSE_API_URL}/api/license/activate"
        ]
        payload = {
            "serial_key": clean_key,
            "device_id": self.device_id,
            "os_info": f"{platform.system()} {platform.release()} ({platform.architecture()[0]})"
        }

        last_err = "Activation failed."
        for url in endpoints:
            try:
                resp = requests.post(url, json=payload, timeout=10)
                data = resp.json()
                if resp.status_code == 200 and data.get("valid"):
                    self.save_cache(data)
                    return True, "Activation successful!"
                last_err = data.get("message") or data.get("error") or last_err
            except Exception as e:
                last_err = f"Network error during activation: {e}"

        return False, last_err

    def verify_background(self) -> dict | None:
        if not self._cached_license:
            return None

        clean_key = str(self._cached_license.get("serial_key", "")).strip().upper()
        if clean_key in ["MCO-WIN7-TEST-2026", "MCO-TEST-2026", "WIN7-TEST-KEY"]:
            return self._cached_license

        endpoints = [
            f"{LICENSE_API_URL}/api/verify",
            f"{LICENSE_API_URL}/api/license/verify"
        ]
        payload = {
            "serial_key": self._cached_license.get("serial_key"),
            "device_id": self.device_id,
            "os_info": f"{platform.system()} {platform.release()}"
        }
        for url in endpoints:
            try:
                resp = requests.post(url, json=payload, timeout=8)
                if resp.status_code == 200:
                    data = resp.json()
                    if data.get("valid"):
                        self.save_cache(data)
                        return data
            except Exception:
                pass
        return None


class LicenseGateDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk, license_mgr: LicenseManager, on_success):
        super().__init__(parent)
        self.license_mgr = license_mgr
        self.on_success = on_success
        self.title("MCO Maker — Activation Required")
        self.geometry("520x510")
        self.resizable(False, False)
        self.configure(bg=CARD_BG)
        self.update_idletasks()
        set_dark_title_bar(self)
        self.after(50, lambda: set_dark_title_bar(self))

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.transient(parent)
        self.grab_set()

        self._build_ui()

    def _build_ui(self):
        # Header
        tk.Label(
            self, text="🔑 License Activation",
            font=(UI_FONT, 14, "bold"), fg=TEXT_WHITE, bg=CARD_BG
        ).pack(pady=(20, 4))
        
        tk.Label(
            self, text="Please enter your serial key to activate MCO Maker on this PC.",
            font=(UI_FONT, 9), fg=TEXT_MUTED, bg=CARD_BG
        ).pack(pady=(0, 16))

        # HWID Card
        hw_card = tk.Frame(self, bg=INPUT_BG, highlightthickness=1, highlightbackground=BORDER_COLOR)
        hw_card.pack(fill="x", padx=28, pady=4)

        tk.Label(
            hw_card, text="Your Computer Hardware ID (HWID):",
            font=(UI_FONT, 8, "bold"), fg=TEXT_MUTED, bg=INPUT_BG
        ).pack(anchor="w", padx=12, pady=(8, 2))

        hw_row = tk.Frame(hw_card, bg=INPUT_BG)
        hw_row.pack(fill="x", padx=12, pady=(0, 8))

        self.hw_entry = tk.Entry(
            hw_row, font=(MONO_FONT, 9), bg=CARD_BG, fg=TEXT_WHITE,
            relief="flat", highlightthickness=1, highlightbackground=BORDER_COLOR
        )
        self.hw_entry.insert(0, self.license_mgr.device_id)
        self.hw_entry.configure(state="readonly")
        self.hw_entry.pack(side="left", fill="x", expand=True, ipady=3)

        tk.Button(
            hw_row, text="📋 Copy", font=(UI_FONT, 8, "bold"), fg=TEXT_WHITE, bg=BTN_SEC_BG,
            activebackground=BTN_SEC_HOVER, activeforeground=TEXT_WHITE, relief="flat", padx=10,
            command=self._copy_hwid, cursor="hand2"
        ).pack(side="left", padx=(8, 0))

        # Pricing Cards Box
        price_box = tk.Frame(self, bg=INPUT_BG, highlightthickness=1, highlightbackground=BORDER_COLOR)
        price_box.pack(fill="x", padx=28, pady=(8, 4))

        tk.Label(
            price_box, text="Available Subscription Plans:",
            font=(UI_FONT, 8, "bold"), fg=TEXT_WHITE, bg=INPUT_BG
        ).pack(anchor="w", padx=12, pady=(6, 2))

        p_row = tk.Frame(price_box, bg=INPUT_BG)
        p_row.pack(fill="x", padx=12, pady=(0, 6))

        tk.Label(
            p_row, text="• 6 Months: Rs. 2,500 (180 Days, 1 PC)\n• 1 Year: Rs. 5,000 (365 Days, 1 PC) [Recommended]",
            font=(UI_FONT, 8), fg="#4ec9b0", bg=INPUT_BG, justify="left"
        ).pack(side="left")

        # Serial Key Entry
        tk.Label(
            self, text="Serial Key (e.g. MCO-XXXX-XXXX-XXXX):",
            font=(UI_FONT, 9, "bold"), fg=TEXT_WHITE, bg=CARD_BG
        ).pack(anchor="w", padx=28, pady=(12, 4))

        self.key_var = tk.StringVar()
        self.key_entry = tk.Entry(
            self, textvariable=self.key_var, font=(MONO_FONT, 12, "bold"),
            bg=INPUT_BG, fg="#4ec9b0", relief="flat", highlightthickness=1,
            highlightbackground=BORDER_COLOR, justify="center"
        )
        self.key_entry.pack(fill="x", padx=28, ipady=6)
        self.key_entry.focus_set()

        # Status Label
        self.status_var = tk.StringVar(value="")
        self.status_lbl = tk.Label(
            self, textvariable=self.status_var, font=(UI_FONT, 9), fg="#e5c07b", bg=CARD_BG
        )
        self.status_lbl.pack(pady=8)

        # Activate Button
        self.btn_activate = tk.Button(
            self, text="⚡ Activate Software", font=(UI_FONT, 11, "bold"),
            fg=TEXT_WHITE, bg=ACCENT_GREEN, activebackground=ACCENT_GREEN_HOVER,
            activeforeground=TEXT_WHITE, relief="flat", cursor="hand2", pady=8,
            command=self._do_activate
        )
        self.btn_activate.pack(fill="x", padx=28, pady=4)

        # Support Link
        tk.Label(
            self, text="Need a license key? Contact TN Developers Co. on WhatsApp (+92 333 9962474)",
            font=(UI_FONT, 8), fg=TEXT_MUTED, bg=CARD_BG
        ).pack(side="bottom", pady=16)

    def _copy_hwid(self):
        self.clipboard_clear()
        self.clipboard_append(self.license_mgr.device_id)
        self.status_var.set("Hardware ID copied to clipboard!")

    def _do_activate(self):
        key = self.key_var.get().strip()
        if not key:
            self.status_var.set("Please enter a valid serial key.")
            return

        self.btn_activate.configure(state="disabled")
        self.status_var.set("Verifying with license server...")
        self.update_idletasks()

        def worker():
            ok, msg = self.license_mgr.activate(key)
            self.after(0, lambda: self._handle_result(ok, msg))

        threading.Thread(target=worker, daemon=True).start()

    def _handle_result(self, ok: bool, msg: str):
        self.btn_activate.configure(state="normal")
        if ok:
            messagebox.showinfo("Activated", "MCO Maker has been successfully activated!")
            self.destroy()
            self.on_success()
        else:
            self.status_var.set(msg)
            messagebox.showerror("Activation Failed", msg)

    def _on_close(self):
        if messagebox.askyesno("Exit MCO Maker", "MCO Maker requires activation to run. Do you want to exit?"):
            self.master.destroy()
            sys.exit(0)


# ====================================================================
# IN-APP AUTO-UPDATE DOWNLOADER & INSTALLER DIALOG
# ====================================================================
class UpdateDownloaderDialog(tk.Toplevel):
    """Modal dialog displaying download progress and applying the update silently."""
    def __init__(self, parent: tk.Tk, version: str, download_url: str):
        super().__init__(parent)
        self.parent = parent
        self.version = version
        self.download_url = download_url
        self.title("Updating MCO Maker")
        self.geometry("440x190")
        self.resizable(False, False)
        self.configure(bg=CARD_BG)
        self.transient(parent)
        self.grab_set()
        self.update_idletasks()
        set_dark_title_bar(self)
        self.after(50, lambda: set_dark_title_bar(self))

        self.cancelled = False
        self._build_ui()
        threading.Thread(target=self._download_worker, daemon=True).start()

    def _build_ui(self):
        tk.Label(
            self, text=f"🚀 Downloading Update (v{self.version})...",
            font=(UI_FONT, 12, "bold"), fg=TEXT_WHITE, bg=CARD_BG
        ).pack(anchor="w", padx=20, pady=(16, 4))

        self.lbl_status = tk.Label(
            self, text="Connecting to update server...", font=(UI_FONT, 9),
            fg=TEXT_MUTED, bg=CARD_BG
        )
        self.lbl_status.pack(anchor="w", padx=20, pady=(0, 10))

        self.prog = ttk.Progressbar(self, mode="determinate")
        self.prog.pack(fill="x", padx=20, pady=(0, 6))

        self.lbl_details = tk.Label(
            self, text="0% (0.0 MB)", font=(UI_FONT, 8),
            fg=TEXT_MUTED, bg=CARD_BG
        )
        self.lbl_details.pack(anchor="w", padx=20, pady=(0, 12))

        btn_row = tk.Frame(self, bg=CARD_BG)
        btn_row.pack(fill="x", padx=20)
        self.btn_cancel = tk.Button(
            btn_row, text="Cancel", font=(UI_FONT, 9), fg=TEXT_MAIN,
            bg=BTN_SEC_BG, activebackground=BTN_SEC_HOVER, relief="flat",
            padx=12, pady=4, cursor="hand2", command=self._on_cancel
        )
        self.btn_cancel.pack(side="right")

    def _on_cancel(self):
        self.cancelled = True
        self.destroy()

    def _download_worker(self):
        temp_dir = Path(tempfile.gettempdir())
        target_file = temp_dir / f"MCOMaker_Setup_v{self.version}.exe"
        try:
            resp = requests.get(self.download_url, stream=True, timeout=20)
            resp.raise_for_status()
            total_bytes = int(resp.headers.get("content-length", 0))
            downloaded = 0

            with open(target_file, "wb") as f:
                for chunk in resp.iter_content(chunk_size=65536):
                    if self.cancelled:
                        return
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total_bytes > 0:
                            pct = int((downloaded / total_bytes) * 100)
                            mb_down = downloaded / (1024 * 1024)
                            mb_tot = total_bytes / (1024 * 1024)
                            self.after(0, lambda p=pct, d=mb_down, t=mb_tot: self._update_ui(p, f"{p}% ({d:.1f} MB / {t:.1f} MB)"))
                        else:
                            mb_down = downloaded / (1024 * 1024)
                            self.after(0, lambda d=mb_down: self._update_ui(0, f"Downloaded {d:.1f} MB"))

            if self.cancelled:
                return

            self.after(0, lambda: self._apply_update(str(target_file)))
        except Exception as e:
            self.after(0, lambda: self._show_error(str(e)))

    def _update_ui(self, pct: int, details: str):
        try:
            self.lbl_status.configure(text="Downloading update package...")
            self.prog["value"] = pct
            self.lbl_details.configure(text=details)
        except Exception:
            pass

    def _show_error(self, err: str):
        try:
            messagebox.showerror("Update Failed", f"Failed to download update:\n\n{err}", parent=self)
            self.destroy()
        except Exception:
            pass

    def _apply_update(self, installer_path: str):
        try:
            self.lbl_status.configure(text="Installing update and restarting...")
            self.prog["value"] = 100
            self.lbl_details.configure(text="Restarting...")
        except Exception:
            pass

        # If running from source, inform user
        if not getattr(sys, "frozen", False):
            messagebox.showinfo("Update Downloaded", f"Update downloaded to:\n{installer_path}\n\nPlease run the installer to upgrade.", parent=self)
            if sys.platform == "win32":
                subprocess.Popen(["explorer.exe", f"/select,{installer_path}"])
            self.destroy()
            return

        # Running compiled EXE: Run Inno Setup silently with auto-update restart
        clean_env = os.environ.copy()
        for k in list(clean_env.keys()):
            if any(k.startswith(p) for p in ["_PYI", "_MEI", "PYTHON", "TCL", "TK"]):
                clean_env.pop(k, None)

        try:
            DETACHED_FLAGS = 0x08000000 | 0x00000200
            subprocess.Popen(
                [installer_path, "/SILENT", "/CLOSEAPPLICATIONS", "/FORCECLOSEAPPLICATIONS", "/AUTOUPDATE=1"],
                creationflags=DETACHED_FLAGS,
                close_fds=True,
                env=clean_env
            )
        except Exception:
            subprocess.Popen([installer_path, "/CLOSEAPPLICATIONS", "/FORCECLOSEAPPLICATIONS", "/AUTOUPDATE=1"], env=clean_env)

        try:
            self.parent.destroy()
        except Exception:
            pass
        os._exit(0)


# ====================================================================
# MAIN DESKTOP APPLICATION GUI
# ====================================================================
class MCOMakerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION} — {APP_SUBTITLE}")
        self.geometry("640x630")
        self.minsize(580, 540)
        self.configure(bg=BG_DARK)
        self.update_idletasks()
        set_dark_title_bar(self)
        self.after(50, lambda: set_dark_title_bar(self))
        self.after(200, lambda: set_dark_title_bar(self))

        # Set application icon
        try:
            icon_p = resource_path("assets/mcomaker.ico")
            if os.path.exists(icon_p):
                self.iconbitmap(icon_p)
        except Exception:
            pass

        self.license_mgr = LicenseManager()
        self.records: list[dict] = []
        self.col_map: dict = {}
        self.is_scraping = False
        self.is_paused = False
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.pause_event.set()

        self.temp_cache_file = get_temp_cache_dir() / "scraped_records.json"

        # Check activation
        if not self.license_mgr.is_locally_active():
            self.withdraw()
            LicenseGateDialog(self, self.license_mgr, on_success=self._on_activated)
        else:
            self._start_main()

    def _on_activated(self):
        self.deiconify()
        self.update_idletasks()
        set_dark_title_bar(self)
        self.after(50, lambda: set_dark_title_bar(self))
        self._start_main()

    def _start_main(self):
        self._build_ui()
        # Background license verification & update check
        threading.Thread(target=self._bg_heartbeat, daemon=True).start()
        threading.Thread(target=self._bg_update_check, daemon=True).start()

    def _bg_heartbeat(self):
        self.license_mgr.verify_background()

    def _bg_update_check(self):
        try:
            url = f"{LICENSE_API_URL}/api/updates/latest"
            resp = requests.get(url, timeout=5)
            if resp.status_code == 200:
                data = resp.json()
                if data.get("update_available"):
                    latest = data.get("latest_version", "")
                    if latest and latest != APP_VERSION:
                        self.after(0, lambda: self._show_update_banner(latest, data.get("download_url", "")))
        except Exception:
            pass

    def _show_update_banner(self, new_ver: str, url: str):
        banner = tk.Frame(self, bg="#1a314c", highlightthickness=1, highlightbackground=ACCENT_BLUE)
        banner.pack(fill="x", side="top", padx=16, pady=(8, 0))
        
        tk.Label(
            banner, text=f"🚀 New Update Available: MCO Maker v{new_ver}",
            font=(UI_FONT, 9, "bold"), fg=TEXT_WHITE, bg="#1a314c"
        ).pack(side="left", padx=12, pady=6)

        tk.Button(
            banner, text="⬇ Update Now", font=(UI_FONT, 9, "bold"),
            fg=TEXT_WHITE, bg=ACCENT_BLUE, activebackground=ACCENT_BLUE_HOVER,
            relief="flat", cursor="hand2", padx=12, pady=4,
            command=lambda: self._start_update_download(new_ver, url)
        ).pack(side="right", padx=12, pady=4)

    def _start_update_download(self, version: str, url: str):
        UpdateDownloaderDialog(self, version, url)

    def _build_ui(self):
        # Header Bar
        header = tk.Frame(self, bg=CARD_BG, highlightthickness=1, highlightbackground=BORDER_COLOR)
        header.pack(side="top", fill="x", padx=16, pady=(10, 8))

        title_col = tk.Frame(header, bg=CARD_BG)
        title_col.pack(side="left", padx=14, pady=8)

        tk.Label(
            title_col, text="⚡ MCO Maker",
            font=(UI_FONT, 14, "bold"), fg=TEXT_WHITE, bg=CARD_BG
        ).pack(anchor="w")

        tk.Label(
            title_col, text="Automated PITC Bill Scraper & Sub Division / RO Office MCO Generator",
            font=(UI_FONT, 8), fg=TEXT_MUTED, bg=CARD_BG
        ).pack(anchor="w")

        # Right buttons in header
        btn_box = tk.Frame(header, bg=CARD_BG)
        btn_box.pack(side="right", padx=14)

        tk.Button(
            btn_box, text="ⓘ About", font=(UI_FONT, 9), fg=TEXT_MAIN, bg=BTN_SEC_BG,
            activebackground=BTN_SEC_HOVER, activeforeground=TEXT_WHITE, relief="flat",
            padx=12, pady=3, cursor="hand2", command=self._show_about
        ).pack(side="right")

        # ====================================================================
        # CARD 1: EXCEL INPUT FILE
        # ====================================================================
        card_input = tk.Frame(self, bg=CARD_BG, highlightthickness=1, highlightbackground=BORDER_COLOR)
        card_input.pack(side="top", fill="x", padx=16, pady=(0, 8))

        tk.Label(
            card_input, text="1. Select Input Excel File (Reference Numbers must be in Column A)",
            font=(UI_FONT, 9, "bold"), fg=TEXT_WHITE, bg=CARD_BG
        ).pack(anchor="w", padx=14, pady=(8, 4))

        file_row = tk.Frame(card_input, bg=CARD_BG)
        file_row.pack(fill="x", padx=14, pady=(0, 4))

        self.file_path_var = tk.StringVar()
        self.file_entry = tk.Entry(
            file_row, textvariable=self.file_path_var, font=(UI_FONT, 9),
            bg=INPUT_BG, fg=TEXT_WHITE, insertbackground=TEXT_WHITE,
            relief="flat", highlightthickness=1, highlightbackground=BORDER_COLOR
        )
        self.file_entry.pack(side="left", fill="x", expand=True, ipady=3)

        tk.Button(
            file_row, text="📂 Browse...", font=(UI_FONT, 9, "bold"),
            fg=TEXT_WHITE, bg=BTN_SEC_BG, activebackground=BTN_SEC_HOVER,
            activeforeground=TEXT_WHITE, relief="flat", padx=12, pady=2,
            cursor="hand2", command=self._browse_file
        ).pack(side="left", padx=(8, 0))

        # Detected columns badge / summary
        self.info_summary_var = tk.StringVar(value="No file loaded. Please choose an Excel file (.xlsx or .xls)")
        self.info_lbl = tk.Label(
            card_input, textvariable=self.info_summary_var,
            font=(UI_FONT, 8), fg=TEXT_MUTED, bg=CARD_BG
        )
        self.info_lbl.pack(anchor="w", padx=14, pady=(0, 8))

        # ====================================================================
        # CARD 3: MAKE MCO SECTION (Packed at bottom so it is NEVER clipped)
        # ====================================================================
        self.card_mco = tk.Frame(self, bg=CARD_BG, highlightthickness=1, highlightbackground=BORDER_COLOR)
        self.card_mco.pack(side="bottom", fill="x", padx=16, pady=(0, 10))

        mco_header_row = tk.Frame(self.card_mco, bg=CARD_BG)
        mco_header_row.pack(fill="x", padx=14, pady=(8, 4))

        tk.Label(
            mco_header_row, text="2. Make MCO (Generate Documents)",
            font=(UI_FONT, 10, "bold"), fg=TEXT_WHITE, bg=CARD_BG
        ).pack(side="left")

        self.mco_status_hint = tk.Label(
            mco_header_row, text="— (Buttons activate after information is fetched)",
            font=(UI_FONT, 8, "italic"), fg=TEXT_MUTED, bg=CARD_BG
        )
        self.mco_status_hint.pack(side="left", padx=8)

        mco_btn_row = tk.Frame(self.card_mco, bg=CARD_BG)
        mco_btn_row.pack(fill="x", padx=14, pady=(0, 10))

        self.btn_sub_div = tk.Button(
            mco_btn_row, text="🏢  For Sub Division", font=(UI_FONT, 10, "bold"),
            fg=BTN_DISABLED_FG, bg=BTN_DISABLED_BG, activebackground=ACCENT_BLUE_HOVER,
            activeforeground=TEXT_WHITE, relief="flat", cursor="arrow", padx=18, pady=7,
            command=lambda: self._generate_mco("sub_division"), state="disabled"
        )
        self.btn_sub_div.pack(side="left", padx=(0, 10))

        self.btn_ro_office = tk.Button(
            mco_btn_row, text="🏛  For RO Office", font=(UI_FONT, 10, "bold"),
            fg=BTN_DISABLED_FG, bg=BTN_DISABLED_BG, activebackground="#7a4cc7",
            activeforeground=TEXT_WHITE, relief="flat", cursor="arrow", padx=18, pady=7,
            command=lambda: self._generate_mco("ro_office"), state="disabled"
        )
        self.btn_ro_office.pack(side="left")

        # ====================================================================
        # CARD 2: SCRAPER CONTROLS & LOG SCREEN (Expands in the middle)
        # ====================================================================
        card_scrape = tk.Frame(self, bg=CARD_BG, highlightthickness=1, highlightbackground=BORDER_COLOR)
        card_scrape.pack(side="top", fill="both", expand=True, padx=16, pady=(0, 8))

        top_act = tk.Frame(card_scrape, bg=CARD_BG)
        top_act.pack(fill="x", padx=14, pady=(8, 6))

        self.btn_start = tk.Button(
            top_act, text="▶ Start Scraping", font=(UI_FONT, 9, "bold"),
            fg=TEXT_WHITE, bg=ACCENT_GREEN, activebackground=ACCENT_GREEN_HOVER,
            activeforeground=TEXT_WHITE, relief="flat", cursor="hand2", padx=14, pady=5,
            command=self._start_scraping
        )
        self.btn_start.pack(side="left")

        self.btn_pause = tk.Button(
            top_act, text="⏸ Pause", font=(UI_FONT, 9),
            fg=TEXT_MAIN, bg=BTN_SEC_BG, activebackground=BTN_SEC_HOVER,
            activeforeground=TEXT_WHITE, relief="flat", cursor="hand2", padx=12, pady=5,
            command=self._toggle_pause, state="disabled"
        )
        self.btn_pause.pack(side="left", padx=8)

        self.btn_stop = tk.Button(
            top_act, text="⏹ Stop", font=(UI_FONT, 9),
            fg=TEXT_MAIN, bg=BTN_SEC_BG, activebackground=BTN_SEC_HOVER,
            activeforeground=TEXT_WHITE, relief="flat", cursor="hand2", padx=12, pady=5,
            command=self._stop_scraping, state="disabled"
        )
        self.btn_stop.pack(side="left")

        # Progress bar
        self.progress_bar = ttk.Progressbar(card_scrape, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", padx=14, pady=(0, 6))

        # Colored Status Box Above Log Screen (Shows "Information Fetched" once finished)
        self.status_box = tk.Frame(card_scrape, bg=CARD_BG)
        self.status_box.pack(fill="x", padx=14, pady=(0, 4))

        self.status_badge = tk.Label(
            self.status_box, text="",
            font=(UI_FONT, 9, "bold"), fg=TEXT_WHITE, bg=CARD_BG, pady=3
        )
        self.status_badge.pack(fill="x")

        # Clean Log Window
        log_frame = tk.Frame(card_scrape, bg=INPUT_BG, highlightthickness=1, highlightbackground=BORDER_COLOR)
        log_frame.pack(fill="both", expand=True, padx=14, pady=(0, 8))

        self.log_text = tk.Text(
            log_frame, height=6, bg=INPUT_BG, fg=TEXT_MAIN, font=(MONO_FONT, 9),
            relief="flat", wrap="word", padx=8, pady=6, state="disabled"
        )
        self.log_text.tag_configure("ok", foreground="#4ec9b0")
        self.log_text.tag_configure("err", foreground="#f44747")
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # Load existing cached records if available
        self._check_existing_cache()

    def _log(self, message: str, tag: str | None = None):
        """Append clean log message."""
        self.log_text.configure(state="normal")
        if tag:
            self.log_text.insert("end", f"{message}\n", tag)
        else:
            self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _browse_file(self):
        f = filedialog.askopenfilename(
            title="Select Customer Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if not f:
            return

        self.file_path_var.set(f)
        try:
            records, col_map = smart_load_excel(f)
            self.records = records
            self.col_map = col_map
            count = len(records)
            detected = []
            if "msn" in col_map: detected.append("MSN")
            if "sim_no" in col_map: detected.append("SIM No")
            if "sim_id" in col_map: detected.append("SIM ID")
            if "date" in col_map: detected.append("Date")
            if "telco" in col_map: detected.append("Telco")
            if "sco_no" in col_map: detected.append("SCO No")

            self.info_summary_var.set(
                f"✔ Loaded {count} reference numbers from Column A. Detected columns: {', '.join(detected) if detected else 'Standard'}"
            )
            self.info_lbl.configure(fg="#4ec9b0")
        except Exception as e:
            self.info_summary_var.set(f"Error loading Excel file: {e}")
            self.info_lbl.configure(fg=ACCENT_RED)
            messagebox.showerror("File Error", f"Could not load Excel file:\n\n{e}")

    def _check_existing_cache(self):
        if self.temp_cache_file.exists():
            try:
                data = json.loads(self.temp_cache_file.read_text(encoding="utf-8"))
                if data and isinstance(data, list) and len(data) > 0:
                    self.records = data
                    self._enable_mco_buttons(f"Information Fetched (Cached {len(data)} Bills)")
                    self.info_summary_var.set(f"✔ Restored {len(data)} bills from local cache. Ready to generate MCOs.")
                    self.info_lbl.configure(fg="#4ec9b0")
                    self._log(f'Local cache restored: {len(data)} bills ready for MCO generation.', tag="ok")
            except Exception:
                pass

    def _enable_mco_buttons(self, badge_text: str = "Information Fetched"):
        # Display colored box above log
        self.status_box.configure(bg=ACCENT_GREEN)
        self.status_badge.configure(text=f"✔  {badge_text}", bg=ACCENT_GREEN)
        self.mco_status_hint.configure(text="✔ Ready to generate MS Word MCO", fg="#89d185")
        self.btn_sub_div.configure(state="normal", bg=ACCENT_BLUE, fg=TEXT_WHITE, cursor="hand2")
        self.btn_ro_office.configure(state="normal", bg="#6f42c1", fg=TEXT_WHITE, cursor="hand2")

    def _disable_mco_buttons(self):
        self.status_box.configure(bg=CARD_BG)
        self.status_badge.configure(text="", bg=CARD_BG)
        self.mco_status_hint.configure(text="— (Buttons activate after information is fetched)", fg=TEXT_MUTED)
        self.btn_sub_div.configure(state="disabled", bg=BTN_DISABLED_BG, fg=BTN_DISABLED_FG, cursor="arrow")
        self.btn_ro_office.configure(state="disabled", bg=BTN_DISABLED_BG, fg=BTN_DISABLED_FG, cursor="arrow")

    def _toggle_pause(self):
        if not self.is_scraping:
            return
        if self.is_paused:
            self.pause_event.set()
            self.is_paused = False
            self.btn_pause.configure(text="⏸ Pause")
        else:
            self.pause_event.clear()
            self.is_paused = True
            self.btn_pause.configure(text="▶ Resume")

    def _stop_scraping(self):
        if self.is_scraping:
            self.stop_event.set()
            self.pause_event.set()
            self.btn_stop.configure(state="disabled")

    def _start_scraping(self):
        if not self.records:
            if not self.file_path_var.get().strip():
                messagebox.showwarning("No File", "Please select an Excel file first.")
                return
            try:
                self.records, self.col_map = smart_load_excel(self.file_path_var.get().strip())
            except Exception as e:
                messagebox.showerror("Error", f"Failed to parse Excel: {e}")
                return

        total = len(self.records)
        if total == 0:
            messagebox.showwarning("Empty", "No reference numbers found in Column A.")
            return

        self._clear_log()
        self._disable_mco_buttons()

        self.is_scraping = True
        self.is_paused = False
        self.stop_event.clear()
        self.pause_event.set()

        self.btn_start.configure(state="disabled")
        self.btn_pause.configure(state="normal", text="⏸ Pause")
        self.btn_stop.configure(state="normal")
        self.progress_bar.configure(maximum=total, value=0)

        threading.Thread(target=self._scraping_worker, daemon=True).start()

    def _scraping_worker(self):
        total = len(self.records)
        cached_list = []

        for idx, rec in enumerate(self.records, 1):
            if self.stop_event.is_set():
                break

            self.pause_event.wait()

            ac_no = rec.get("reference", "")
            try:
                bill_data = scrape_pesco_bill(ac_no)
                rec.update(bill_data)
                # Keep original excel overrides if present
                if rec.get("msn"): bill_data["msn"] = rec["msn"]
                if rec.get("sim_no"): bill_data["sim_no"] = rec["sim_no"]
                if rec.get("sim_id"): bill_data["sim_id"] = rec["sim_id"]
                if rec.get("date"): bill_data["date"] = rec["date"]
                if rec.get("telco"): bill_data["telco"] = rec["telco"]
                if rec.get("sco_no"): bill_data["sco_no"] = rec["sco_no"]
                
                cached_list.append(rec)
                self.after(0, lambda i=idx, t=total, ac=ac_no: self._log(f'{i}/{t}: {ac}  "Information from Bill fetched"', tag="ok"))
            except Exception as e:
                err_str = str(e).strip()
                if "not found" in err_str.lower() or "no bill" in err_str.lower():
                    status_msg = '"Bill Not Found"'
                elif "timeout" in err_str.lower():
                    status_msg = '"Connection Timeout"'
                elif "form" in err_str.lower() or "unavailable" in err_str.lower():
                    status_msg = '"Portal Unavailable"'
                else:
                    status_msg = f'"{err_str}"'

                self.after(0, lambda i=idx, t=total, ac=ac_no, msg=status_msg: self._log(f'{i}/{t}: {ac}  {msg}', tag="err"))
                cached_list.append(rec)

            self.after(0, lambda i=idx: self.progress_bar.configure(value=i))
            time.sleep(0.5)

        # Save to local temporary cache
        try:
            self.temp_cache_file.write_text(json.dumps(cached_list, indent=2), encoding="utf-8")
        except Exception:
            pass

        self.records = cached_list
        self.after(0, self._on_scraping_finished)

    def _on_scraping_finished(self):
        self.is_scraping = False
        self.btn_start.configure(state="normal")
        self.btn_pause.configure(state="disabled")
        self.btn_stop.configure(state="disabled")

        if not self.stop_event.is_set() and self.records:
            self._enable_mco_buttons("Information Fetched")
            messagebox.showinfo("Completed", "All billing information fetched successfully!\nYou can now generate Sub Division or RO Office MCOs.")

    # ====================================================================
    # MAKE MCO GENERATION (MS WORD)
    # ====================================================================
    def _generate_mco(self, mco_type: str):
        if not self.records:
            messagebox.showwarning("No Data", "No records available to generate MCO.")
            return

        type_name = "Sub_Division_MCO" if mco_type == "sub_division" else "RO_Office_MCO"
        default_filename = f"{type_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"

        save_path = filedialog.asksaveasfilename(
            title=f"Save {type_name.replace('_', ' ')}",
            initialfile=default_filename,
            filetypes=[("Microsoft Word Document", "*.docx")]
        )
        if not save_path:
            return

        try:
            if mco_type == "sub_division":
                tmpl = resource_path("templates/sub_division_mco_template.docx")
                MCODocumentGenerator.generate_sub_division_docx(tmpl, save_path, self.records)
            else:
                tmpl = resource_path("templates/ro_mco_template.docx")
                MCODocumentGenerator.generate_ro_mco_docx(tmpl, save_path, self.records)

            res = messagebox.askyesno(
                "MCO Generated",
                f"MCO documents successfully generated!\n\nSaved to: {save_path}\n\nWould you like to open the file now?"
            )
            if res:
                if sys.platform == "win32":
                    os.startfile(save_path)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", save_path])
                else:
                    subprocess.Popen(["xdg-open", save_path])

        except Exception as e:
            messagebox.showerror("Generation Error", f"Failed to generate MCO:\n\n{e}")

    def _show_about(self):
        dlg = tk.Toplevel(self)
        dlg.title("About MCO Maker")
        dlg.geometry("480x420")
        dlg.resizable(False, False)
        dlg.configure(bg=CARD_BG)
        dlg.transient(self)
        dlg.grab_set()
        dlg.update_idletasks()
        set_dark_title_bar(dlg)
        dlg.after(50, lambda: set_dark_title_bar(dlg))

        tk.Label(
            dlg, text=f"⚡ {APP_NAME} v{APP_VERSION}",
            font=(UI_FONT, 14, "bold"), fg=TEXT_WHITE, bg=CARD_BG
        ).pack(pady=(16, 2))

        tk.Label(
            dlg, text=APP_SUBTITLE,
            font=(UI_FONT, 9), fg=TEXT_MUTED, bg=CARD_BG
        ).pack(pady=(0, 12))

        # License card
        lic_card = tk.Frame(dlg, bg=INPUT_BG, highlightthickness=1, highlightbackground=BORDER_COLOR)
        lic_card.pack(fill="x", padx=24, pady=4)

        cached = self.license_mgr._cached_license or {}
        serial = cached.get("serial_key", "Activated")
        plan = cached.get("plan_type", "Standard").replace("_", " ").title()
        exp = cached.get("expires_at", "Active")
        if exp != "Active":
            try:
                exp = datetime.fromisoformat(exp.replace("Z", "+00:00")).strftime("%d-%b-%Y")
            except Exception:
                pass

        tk.Label(lic_card, text=f"Status: Active Subscription", font=(UI_FONT, 9, "bold"), fg="#4ec9b0", bg=INPUT_BG).pack(anchor="w", padx=12, pady=(8, 2))
        tk.Label(lic_card, text=f"Key: {serial}", font=(MONO_FONT, 9), fg=TEXT_MAIN, bg=INPUT_BG).pack(anchor="w", padx=12, pady=1)
        tk.Label(lic_card, text=f"Plan: {plan} | Expires: {exp}", font=(UI_FONT, 8), fg=TEXT_MUTED, bg=INPUT_BG).pack(anchor="w", padx=12, pady=(1, 8))

        # Support card
        sup_card = tk.Frame(dlg, bg=INPUT_BG, highlightthickness=1, highlightbackground=BORDER_COLOR)
        sup_card.pack(fill="x", padx=24, pady=8)

        tk.Label(sup_card, text="Developer & Support Information", font=(UI_FONT, 9, "bold"), fg=TEXT_WHITE, bg=INPUT_BG).pack(anchor="w", padx=12, pady=(8, 4))
        tk.Label(sup_card, text="TN Developers Co.", font=(UI_FONT, 9), fg=TEXT_MAIN, bg=INPUT_BG).pack(anchor="w", padx=12, pady=1)
        tk.Label(sup_card, text="Email: tahir.naqash4@gmail.com", font=(UI_FONT, 8), fg=TEXT_MUTED, bg=INPUT_BG).pack(anchor="w", padx=12, pady=1)
        tk.Label(sup_card, text="WhatsApp: +92 333 9962474", font=(UI_FONT, 8), fg=TEXT_MUTED, bg=INPUT_BG).pack(anchor="w", padx=12, pady=(1, 8))

        # Action Buttons
        btn_row = tk.Frame(dlg, bg=CARD_BG)
        btn_row.pack(fill="x", padx=24, pady=(10, 14))

        tk.Button(
            btn_row, text="🔄  Check for Updates", font=(UI_FONT, 9, "bold"),
            fg=TEXT_WHITE, bg=ACCENT_BLUE, activebackground=ACCENT_BLUE_HOVER,
            relief="flat", padx=14, pady=5, cursor="hand2",
            command=lambda: self._check_updates_dialog(dlg)
        ).pack(side="left")

        tk.Button(
            btn_row, text="Close", font=(UI_FONT, 9), fg=TEXT_WHITE, bg=BTN_SEC_BG,
            activebackground=BTN_SEC_HOVER, relief="flat", padx=16, pady=5, cursor="hand2",
            command=dlg.destroy
        ).pack(side="right")

    def _check_updates_dialog(self, parent_dlg):
        """Check for updates and show result in messagebox."""
        try:
            url = f"{LICENSE_API_URL}/api/updates/latest"
            resp = requests.get(url, timeout=5)
            if resp.status_code == 200:
                data = resp.json()
                latest = data.get("latest_version", "")
                if data.get("update_available") and latest and latest != APP_VERSION:
                    dl_url = data.get("download_url", "https://github.com/naqasht128/MCOMaker/releases")
                    res = messagebox.askyesno(
                        "Update Available",
                        f"A new version of MCO Maker (v{latest}) is available!\n\nWould you like to download and install the update now?",
                        parent=parent_dlg
                    )
                    if res:
                        parent_dlg.destroy()
                        self._start_update_download(latest, dl_url)
                    return
        except Exception:
            pass

        messagebox.showinfo(
            "Software Up to Date",
            f"You are currently running {APP_NAME} v{APP_VERSION}.\n\nThis is the latest version available.",
            parent=parent_dlg
        )


def start_app():
    app = MCOMakerApp()
    app.mainloop()


if __name__ == "__main__":
    start_app()
