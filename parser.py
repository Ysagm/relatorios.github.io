"""
Le planilhas Excel de manutencao aeronautica e gera data.json para o dashboard.

Replica fielmente a logica de parsing do dashboard.html (parseWorkbook em JS).
"""

import json
import os
import re
import unicodedata
from datetime import datetime, date
from pathlib import Path

import xlrd  # leitura de .xls
from openpyxl import load_workbook  # leitura de .xlsx

# ── CONFIG ──────────────────────────────────────────────────────────────────

# Pasta base onde estao as subpastas de cada aeronave. Pode ser sobrescrita
# pela variavel de ambiente ONEDRIVE_BASE (usado no workflow do GitHub Actions,
# que sincroniza o OneDrive para um diretorio do runner).
DEFAULT_BASE = r"C:\Users\ysado\OneDrive"
BASE_DIR = Path(os.environ.get("ONEDRIVE_BASE", DEFAULT_BASE))

# Cada aeronave tem uma subpasta; dentro dela pegamos o arquivo .xls/.xlsx
# mais recente (por data de modificacao), pois o nome do arquivo muda a cada
# atualizacao da planilha (ex: "... - 14-06-2026.xls").
AIRCRAFT_FOLDERS = [
    "Planilhas 2026 - PP-AGN",
    "Planilhas 2026 - PP-VEL",
    "Planilhas 2026 - PS-FLC",
    "Planilhas 2026 - PS-NFA",
]


def find_latest_spreadsheet(folder: Path):
    candidates = [
        f for f in folder.glob("*")
        if f.suffix.lower() in (".xls", ".xlsx") and not f.name.startswith("~$")
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda f: f.stat().st_mtime)

SITE_DIR = Path(__file__).resolve().parent / "site"
OUTPUT_PATH = SITE_DIR / "data.json"
OUTPUT_JS_PATH = SITE_DIR / "data.js"

TARGET_SHEETS = ["Manutenção", "Manutenao", "Componentes", "DIR", "DIR MOTOR", "DIR APU"]

ACFT_NAME_RE = re.compile(r"\b([A-Z]{2}-[A-Z0-9]{3})\b", re.IGNORECASE)


# ── HELPERS ──────────────────────────────────────────────────────────────────

def normalize(v):
    s = "" if v is None else str(v)
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.strip().upper()


def is_alert(v):
    n = normalize(v)
    return n in ("ATENCAO", "ATTENTION")


def excel_date_to_py(serial):
    if serial is None or not isinstance(serial, (int, float)):
        return None
    try:
        from datetime import timedelta
        return date(1899, 12, 30) + timedelta(days=serial)
    except Exception:
        return None


def to_date(val):
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, (int, float)):
        return excel_date_to_py(val)
    if isinstance(val, str):
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                return datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                continue
        return None
    return None


def date_diff_days(val):
    d = to_date(val)
    if d is None:
        return None
    today = date.today()
    return (d - today).days


def extract_acft_name(filename):
    base = Path(filename).stem
    m = ACFT_NAME_RE.search(base)
    if m:
        return m.group(1).upper()
    b = re.search(r"__([A-Z0-9]{5})__", base, re.IGNORECASE)
    if b:
        c = b.group(1).upper()
        return c[:2] + "-" + c[2:]
    return re.split(r"[_\-\s]+", base)[0].upper()


def num(v):
    try:
        if v is None or v == "" or v == "-":
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


# ── LEITURA DE WORKBOOK (abstrai .xls vs .xlsx) ─────────────────────────────

class Sheet:
    def __init__(self, name, rows):
        self.name = name
        self.rows = rows  # list of list, header:1 style (0-indexed rows/cols)


def read_workbook(path):
    """Retorna lista de Sheet com linhas em formato lista-de-listas (defval=None)."""
    ext = Path(path).suffix.lower()
    sheets = []
    if ext == ".xls":
        wb = xlrd.open_workbook(path)
        for sn in wb.sheet_names():
            ws = wb.sheet_by_name(sn)
            rows = []
            for r in range(ws.nrows):
                row = []
                for c in range(ws.ncols):
                    cell = ws.cell(r, c)
                    val = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            tup = xlrd.xldate_as_tuple(val, wb.datemode)
                            val = datetime(*tup)
                        except Exception:
                            pass
                    elif val == "":
                        val = None
                    row.append(val)
                rows.append(row)
            sheets.append(Sheet(sn, rows))
    else:
        wb = load_workbook(path, data_only=True)
        for sn in wb.sheetnames:
            ws = wb[sn]
            rows = []
            for row_cells in ws.iter_rows(values_only=True):
                rows.append(list(row_cells))
            sheets.append(Sheet(sn, rows))
    return sheets


def get(row, idx):
    if idx < len(row):
        return row[idx]
    return None


# ── PARSER (espelha parseWorkbook do JS) ────────────────────────────────────

def parse_workbook(sheets, acft_name):
    tasks = []
    info = {}

    for sheet in sheets:
        sn = sheet.name.strip()
        if not any(sn == t or sn.startswith(t) for t in TARGET_SHEETS):
            continue

        rows = sheet.rows

        # Cabecalho da aeronave
        if "totalHours" not in info:
            for i in range(min(15, len(rows))):
                row = rows[i] or []
                for j in range(len(row) - 1):
                    v = normalize(get(row, j))
                    if v == "HORAS TOTAIS":
                        n = num(get(row, j + 1))
                        if n is not None:
                            info["totalHours"] = n
                    if v == "POUSOS TOTAIS":
                        n = num(get(row, j + 1))
                        if n is not None:
                            info["totalLandings"] = n
                    if v == "CICLOS TOTAIS":
                        n = num(get(row, j + 1))
                        if n is not None:
                            info["totalCycles"] = n
                    if "registration" not in info:
                        cell_str = str(get(row, j) or "")
                        rm = re.match(r"^([A-Z]{2}-[A-Z0-9]{3})$", cell_str, re.IGNORECASE)
                        if rm:
                            info["registration"] = rm.group(1).upper()

        is_dir = sn.startswith("DIR")

        if is_dir:
            hdr = -1
            for i, row in enumerate(rows):
                rs = "|".join(normalize(v) for v in (row or []))
                if ("AIRWORTHINESS" in rs) or ("DIRETRIZ" in rs) or ("AD/DA" in rs) or ("AD DA" in rs):
                    hdr = i
                    break
            if hdr < 0:
                continue
            for i in range(hdr + 1, len(rows)):
                row = rows[i] or []
                alert_h = is_alert(get(row, 14))
                alert_d = is_alert(get(row, 15))
                if not alert_h and not alert_d:
                    continue
                ad = str(get(row, 1) or "").strip()
                sb = str(get(row, 2) or "").strip()
                desc = str(get(row, 3) or "").strip()
                if not ad and not sb and not desc:
                    continue

                alert_types = []
                due_hours_str = None
                due_days_str = None

                if alert_h and get(row, 12) is not None:
                    h = num(get(row, 12))
                    if h is not None:
                        due_hours_str = f"{h:.1f}h"
                        alert_types.append("hours")

                if alert_d:
                    days = date_diff_days(get(row, 13))
                    if days is not None:
                        due_days_str = f"{days} dias"
                        alert_types.append("days")
                    elif get(row, 13) is not None:
                        n = num(get(row, 13))
                        if n is not None:
                            due_days_str = f"{round(n)} dias"
                            alert_types.append("days")

                tasks.append({
                    "id": str(get(row, 0) or "").strip(),
                    "task": ad or sb or "",
                    "description": desc or ad or "AD",
                    "pn": sb or "",
                    "sheet": sn,
                    "dueHoursStr": due_hours_str,
                    "dueDaysStr": due_days_str,
                    "dueCyclesStr": None,
                    "alertTypes": alert_types,
                })
        else:
            hdr = -1
            for i, row in enumerate(rows):
                rs = "|".join(normalize(v) for v in (row or []))
                if ("TASK" in rs) or ("ID" in rs and ("INSPECTIONS" in rs or "NOMENCLATURA" in rs)):
                    hdr = i
                    break
            if hdr < 0:
                continue

            alert_col = 19 if sn == "Componentes" else 18

            for i in range(hdr + 1, len(rows)):
                row = rows[i] or []
                if not is_alert(get(row, alert_col)):
                    continue
                task_id = str(get(row, 1) or "").strip()
                desc = str(get(row, 2) or "").strip()
                if not task_id and not desc:
                    continue

                alert_types = []
                due_hours_str = None
                due_days_str = None
                due_cycles_str = None

                # Horas: col[14] direto (saldo calculado pelo Excel)
                saldo_h = get(row, 14)
                if saldo_h is not None and saldo_h != "-" and saldo_h != "":
                    h = num(saldo_h)
                    if h is not None:
                        due_hours_str = f"{h:.1f}h"
                        alert_types.append("hours")

                # Dias: recalcular com col[13] (Date)
                next_date = get(row, 13)
                if next_date is not None and next_date != "-":
                    days = date_diff_days(next_date)
                    if days is not None:
                        due_days_str = f"{days} dias"
                        alert_types.append("days")
                    else:
                        sd = get(row, 16)
                        if sd is not None and sd != "-":
                            n = num(sd)
                            if n is not None:
                                due_days_str = f"{round(n)} dias"
                                alert_types.append("days")
                else:
                    sd = get(row, 16)
                    if sd is not None and sd != "-" and sd != "":
                        n = num(sd)
                        if n is not None:
                            due_days_str = f"{round(n)} dias"
                            alert_types.append("days")

                # Ciclos: col[15]
                saldo_c = get(row, 15)
                if saldo_c is not None and saldo_c != "-" and saldo_c != "":
                    c = num(saldo_c)
                    if c is not None:
                        due_cycles_str = f"{round(c)} ciclos"
                        alert_types.append("cycles")

                if not alert_types:
                    alert_types.append("days")

                tasks.append({
                    "id": str(get(row, 0) or "").strip(),
                    "task": task_id,
                    "description": desc,
                    "pn": str(get(row, 3) or "").strip(),
                    "sheet": sn,
                    "dueHoursStr": due_hours_str,
                    "dueDaysStr": due_days_str,
                    "dueCyclesStr": due_cycles_str,
                    "alertTypes": alert_types,
                })

    info.setdefault("totalHours", None)
    info.setdefault("totalLandings", None)
    info.setdefault("totalCycles", None)

    return {"tasks": tasks, "info": info, "name": acft_name}


# ── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    aircraft = {}

    for folder_name in AIRCRAFT_FOLDERS:
        folder = BASE_DIR / folder_name
        if not folder.exists():
            print(f"[AVISO] pasta nao encontrada, ignorando: {folder}")
            continue

        p = find_latest_spreadsheet(folder)
        if p is None:
            print(f"[AVISO] nenhuma planilha encontrada em: {folder}")
            continue

        acft_name = extract_acft_name(p.name)
        print(f"Lendo {p.name} -> {acft_name}")
        try:
            sheets = read_workbook(str(p))
            data = parse_workbook(sheets, acft_name)
            aircraft[acft_name] = {"info": data["info"], "tasks": data["tasks"]}
            print(f"  {len(data['tasks'])} alerta(s) encontrado(s)")
        except Exception as e:
            print(f"[ERRO] falha ao processar {p.name}: {e}")

    output = {
        "generated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "aircraft": aircraft,
    }

    json_text = json.dumps(output, ensure_ascii=False, indent=2)
    OUTPUT_PATH.write_text(json_text, encoding="utf-8")
    print(f"\nGerado: {OUTPUT_PATH}")

    OUTPUT_JS_PATH.write_text(
        f"window.EMBEDDED_DATA = {json_text};\n", encoding="utf-8"
    )
    print(f"Gerado: {OUTPUT_JS_PATH}")


if __name__ == "__main__":
    main()
