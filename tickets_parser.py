#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tickets_parser.py
-----------------
Extrae campos de PDFs de tickets (formato Service Desk similar al que enviaste) y
genera un Excel en español con columnas útiles + tiempos calculados.

Uso:
  python tickets_parser.py --pdf_dir "C:/ruta/a/pdfs" --out "C:/ruta/salida/Tickets.xlsx"

Opcionales:
  --pattern "Ticket-*.pdf"   (por defecto procesa TODOS los .pdf del directorio: "*.pdf")
  --timezone "America/Asuncion"
  --dedupe                   (desduplicar por N° Ticket quedando con la fila más completa)
  --dump_text                (guardar debug_texts con el texto leído)

Requisitos:
  pip install -U pandas PyMuPDF PyPDF2 openpyxl
"""
import os, re, sys, argparse, time
import pandas as pd
from datetime import datetime, timedelta
from typing import Dict, List, Optional
from zoneinfo import ZoneInfo
from glob import glob

try:
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.workbook.defined_name import DefinedName
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False
    get_column_letter = None  # type: ignore[assignment]
    DataValidation = None  # type: ignore[assignment]
    DefinedName = None  # type: ignore[assignment]

CBSA_DEPARTMENTS: List[str] = [
    "Mesa de Dinero",
    "IT",
    "Comercial",
    "Mesa de Ayuda",
    "Operaciones",
]

AREA_CBSA = "CBSA"
AREA_FONDOS = "Fondos"
AREA_OPTIONS = [AREA_CBSA, AREA_FONDOS]

# ============== Backends de extracción de texto ==============
try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except Exception:
    HAS_PYMUPDF = False

try:
    from PyPDF2 import PdfReader
    HAS_PYPDF2 = True
except Exception:
    HAS_PYPDF2 = False

def log(msg: str):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)

def extract_text_pymupdf(pdf_path: str) -> str:
    doc = fitz.open(pdf_path)
    chunks = []
    for i, page in enumerate(doc, start=1):
        try:
            chunks.append(page.get_text("text") or "")
        except Exception as e:
            log(f"   - PyMuPDF page {i} error: {e}")
    return "\n".join(chunks).strip()

def extract_text_pypdf2(pdf_path: str) -> str:
    reader = PdfReader(pdf_path)
    chunks = []
    for i, page in enumerate(reader.pages, start=1):
        try:
            txt = page.extract_text() or ""
            chunks.append(txt)
        except Exception as e:
            log(f"   - PyPDF2 page {i} error: {e}")
    return "\n".join(chunks).strip()

def extract_text(pdf_path: str) -> str:
    if HAS_PYMUPDF:
        txt = extract_text_pymupdf(pdf_path)
        if len(txt) > 80:
            return txt
    if HAS_PYPDF2:
        txt = extract_text_pypdf2(pdf_path)
        if len(txt) > 80:
            return txt
    return ""

# ============== Limpieza y patrones ==============
FOOTER_RE = re.compile(r"Ticket\s*#\d+\s+printed by.*?Page\s*\d+", re.I | re.S)
WEEKDAY_LINE = re.compile(r"^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b.*$", re.I | re.M)

STAMP_RE = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{2,4}),?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?", re.I)
TS_LINE_RE = STAMP_RE

HEADER_BLOCK_RE = re.compile(
    r"Ticket\s*#(?P<ticket>\d+).*?Status\s+(?P<Status>.+?)\s+Name\s+(?P<Name>.+?)\s+Priority\s+(?P<Priority>.+?)\s+Email\s+(?P<Email>.+?)\s+Department\s+(?P<Department>.+?)(?=\s+Phone\b|\s+Create Date)(?:\s+Phone\s+.*?\s+)?Create Date\s+(?P<CreateDate>.+?)\s+Source\s+(?P<Source>.+?)\s+Ticket Details",
    re.S | re.I
)

LABELY = re.compile(r"^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ/ .]{2,60}\s*:\s*")

def clean_text(text: str) -> str:
    text = FOOTER_RE.sub("", text)
    text = WEEKDAY_LINE.sub("", text)
    return text

def find_line_idx(lines: List[str], predicate, start=0, limit=None) -> Optional[int]:
    end = len(lines) if limit is None else min(len(lines), start+limit)
    for i in range(start, end):
        if predicate(lines[i]):
            return i
    return None

def is_related_label(s: str) -> bool:
    t = " ".join(s.strip().lower().split())
    return t.endswith(":") and (
        "related client number or user" in t or "related client" in t or
        "cliente relacionado" in t or "usuario relacionado" in t or
        "cliente/usuario relacionado" in t
    )

def extract_title_after_urgency(cleaned: str) -> str:
    lines = cleaned.splitlines()
    td = find_line_idx(lines, lambda ln: ln.strip().lower().startswith("ticket details"), 0, 300)
    start = (td + 1) if td is not None else 0

    urg = find_line_idx(lines, lambda ln: re.match(r"^\s*Urgency", ln.strip(), re.I), start, 200)
    if urg is None:
        return ""

    i = urg + 2  # saltar "Urgency:" y su valor

    if i < len(lines) and is_related_label(lines[i]):
        i += 2

    while i < len(lines):
        s = lines[i].strip()
        if s and not TS_LINE_RE.search(s):
            return s
        i += 1
    return ""

def is_auto(author: str, body: str) -> bool:
    blob = (author + " " + body).lower()
    auto_keys = ["mail delivery subsystem","delivery status notification","mailer-daemon","postmaster","do not reply","no-reply"]
    return any(k in blob for k in auto_keys)

def extract_thread_entries(full_text: str) -> List[Dict[str,str]]:
    entries = []
    matches = list(TS_LINE_RE.finditer(full_text))
    for i, m in enumerate(matches):
        line_end   = full_text.find("\n", m.end())
        if line_end == -1:
            line_end = len(full_text)
        header_line = full_text[m.start():line_end]
        author = header_line[m.end()-m.start():].strip()
        body = full_text[line_end:matches[i+1].start()] if i+1 < len(matches) else full_text[line_end:]
        entries.append({"stamp": m.group(0), "author": author, "body": body.strip()})
    return entries

def fallback_last_author_from_tail(cleaned: str) -> str:
    lines = [ln.strip() for ln in cleaned.splitlines()]
    for s in reversed(lines):
        if not s:
            continue
        if TS_LINE_RE.search(s):
            continue
        if LABELY.match(s):
            continue
        if s.lower() in ("ticket details","urgency:"):
            continue
        return s
    return ""

def parse_timestamp_ddmm(s: str, tz: ZoneInfo) -> Optional[datetime]:
    m = STAMP_RE.search(s or "")
    if not m:
        return None
    d1, d2, yyyy, hh, mm, ss, ampm = m.groups()
    d, mo = int(d1), int(d2)
    year = int(yyyy) + (2000 if int(yyyy) < 100 else 0)
    hour, minute = int(hh), int(mm)
    sec = int(ss) if ss else 0
    if ampm:
        if ampm.upper() == "PM" and hour != 12: hour += 12
        if ampm.upper() == "AM" and hour == 12: hour = 0
    try:
        return datetime(year, mo, d, hour, minute, sec, tzinfo=tz)
    except Exception:
        return None

def parse_timestamp_mmdd(s: str, tz: ZoneInfo) -> Optional[datetime]:
    m = STAMP_RE.search(s or "")
    if not m:
        return None
    d1, d2, yyyy, hh, mm, ss, ampm = m.groups()
    mo, d = int(d1), int(d2)
    year = int(yyyy) + (2000 if int(yyyy) < 100 else 0)
    hour, minute = int(hh), int(mm)
    sec = int(ss) if ss else 0
    if ampm:
        if ampm.upper() == "PM" and hour != 12: hour += 12
        if ampm.upper() == "AM" and hour == 12: hour = 0
    try:
        return datetime(year, mo, d, hour, minute, sec, tzinfo=tz)
    except Exception:
        return None

MAX_FUTURE_DELTA = timedelta(days=365)


def parse_pdf_timestamp(s: str, now: datetime, tz: ZoneInfo) -> Optional[datetime]:
    candidates = []
    for parse_fn in (parse_timestamp_ddmm, parse_timestamp_mmdd):
        dt = parse_fn(s, tz)
        if not dt:
            continue
        if dt > now + MAX_FUTURE_DELTA:
            continue
        candidates.append(dt)

    if not candidates:
        return None

    return min(candidates, key=lambda dt: abs((dt - now).total_seconds()))

def is_open_status(status: str) -> bool:
    if not isinstance(status, str):
        return True
    st = status.strip().lower()
    closed_words = ["closed","resolved","solved","cerrado","resuelto","completado","finalizado","finalizada"]
    return not any(w in st for w in closed_words)

def parse_header_block(cleaned: str, lines: List[str]) -> Dict[str,str]:
    m = HEADER_BLOCK_RE.search(cleaned)
    if m:
        gd = m.groupdict()
        return {
            "Ticket Number": gd.get("ticket","").strip(),
            "Status":        gd.get("Status","").strip(),
            "Name":          gd.get("Name","").strip(),
            "Priority":      gd.get("Priority","").strip(),
            "Department":    gd.get("Department","").strip(),
            "Create Date":   gd.get("CreateDate","").strip(),
        }
    return {"Ticket Number":"","Status":"","Name":"","Priority":"","Department":"","Create Date":""}

def parse_pdf(pdf_path: str, tz: ZoneInfo, now: datetime) -> Dict[str,str]:
    raw = extract_text(pdf_path)
    if not raw or len(raw) < 80:
        return {"N° Ticket":"","Título del ticket":"","Estado BW":"","Prioridad":"","Departamento":"","Fecha de creación":"","Última respuesta por":"","Última respuesta el":"","Error":"no_text_extracted"}
    cleaned = clean_text(raw)
    lines = cleaned.splitlines()
    hdr = parse_header_block(cleaned, lines)
    title = extract_title_after_urgency(cleaned)

    entries = extract_thread_entries(cleaned)
    last_by, last_at = "", ""
    first_author = ""
    for e in entries:
        if e["author"] and not is_auto(e["author"], e["body"]):
            if not first_author:
                first_author = e["author"]
            last_by, last_at = e["author"], e["stamp"]

    if not last_by:
        last_by = fallback_last_author_from_tail(cleaned)

    if not first_author:
        first_author = fallback_last_author_from_tail(cleaned)

    return {
        "N° Ticket": hdr.get("Ticket Number",""),
        "Título del ticket": title,
        "Estado BW": hdr.get("Status",""),
        "Prioridad": hdr.get("Priority",""),
        "Departamento": "",
        "Fecha de creación": hdr.get("Create Date",""),
        "Autor": first_author,
        "Última respuesta por": last_by,
        "Última respuesta el": last_at,
        "Error": ""
    }

def main():
    ap = argparse.ArgumentParser(description="PDF tickets -> Excel (ES)")
    ap.add_argument("--pdf_dir", required=True, help="Carpeta con PDFs")
    ap.add_argument("--out", required=True, help="Ruta de salida del Excel")
    ap.add_argument("--pattern", default="*.pdf", help="Patrón de archivos (glob)")
    ap.add_argument("--timezone", default="America/Asuncion", help="Zona horaria")
    args = ap.parse_args()

    tz = ZoneInfo(args.timezone)
    now = datetime.now(tz)

    files = sorted(glob(os.path.join(args.pdf_dir, args.pattern)))
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files:
        log("No se encontraron PDFs.")
        sys.exit(1)

    log(f"Procesando {len(files)} PDF(s)…")
    rows = []
    for path in files:
        row = parse_pdf(path, tz, now)
        rows.append(row)

    df = pd.DataFrame(rows)

    rename_map = {}
    if "N° Ticket" in df.columns:
        rename_map["N° Ticket"] = "N Ticket"
    if "Estado" in df.columns:
        rename_map["Estado"] = "Estado BW"
    if rename_map:
        df = df.rename(columns=rename_map)

    estado_bw_series = None
    if "Estado BW" in df.columns:
        estado_bw_series = df["Estado BW"].copy()
        df = df.drop(columns=["Estado BW"])

    if "Área" not in df.columns:
        df["Área"] = ""
    else:
        df["Área"] = df["Área"].fillna("").astype(str).str.strip()

    if "Departamento" not in df.columns:
        df["Departamento"] = ""
    else:
        df["Departamento"] = df["Departamento"].fillna("").astype(str).str.strip()

    def normalize_department(row):
        area = (row.get("Área") or "").strip()
        dept = (row.get("Departamento") or "").strip()
        if not dept:
            return ""
        if area == AREA_FONDOS:
            return "-" if dept == "-" else ""
        if area == AREA_CBSA:
            return dept if dept in CBSA_DEPARTMENTS else ""
        return dept

    df["Departamento"] = df.apply(normalize_department, axis=1)

    desired_order = [
        "N Ticket",
        "Título del ticket",
        "Autor",
        "Prioridad",
        "Área",
        "Departamento",
    ]
    leading_cols = [col for col in desired_order if col in df.columns]
    remaining_cols = [col for col in df.columns if col not in leading_cols]
    df = df.loc[:, leading_cols + remaining_cols]

    def human_delta(td_seconds: Optional[float]) -> str:
        if td_seconds is None: return ""
        secs = int(td_seconds)
        if secs <= 0: return ""
        mins, _ = divmod(secs, 60)
        hrs, m = divmod(mins, 60)
        days, h = divmod(hrs, 24)
        if days > 0: return f"{days}d {h}h"
        if hrs > 0: return f"{hrs}h {m}m"
        return f"{m}m"

    def parse_date(s: str) -> Optional[datetime]:
        return parse_pdf_timestamp(s, now, tz)

    df["Tiempo parado desde la última respuesta"] = df["Última respuesta el"].apply(lambda s: human_delta((now - parse_date(s)).total_seconds() if parse_date(s) else None))

    def compute_open_age(row):
        st = ""
        if estado_bw_series is not None:
            st = estado_bw_series.get(row.name, "")
        cd = parse_date(row.get("Fecha de creación",""))
        if cd and is_open_status(st):
            return human_delta((now - cd).total_seconds())
        return ""
    df["Tiempo abierto (si sigue abierto)"] = df.apply(compute_open_age, axis=1)

    out_path = args.out
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)

    if not HAS_OPENPYXL:
        raise RuntimeError(
            "Se requiere openpyxl para generar el Excel con listas desplegables. "
            "Instalá openpyxl (pip install openpyxl) e intentá nuevamente."
        )

    sheet_name = "Tickets"
    data_row_start = 2
    data_row_end = max(len(df) + 1, 200)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        wb = writer.book
        ws = writer.sheets[sheet_name]

        area_idx = df.columns.get_loc("Área") + 1
        dept_idx = df.columns.get_loc("Departamento") + 1
        area_col = get_column_letter(area_idx)
        dept_col = get_column_letter(dept_idx)
        area_range = f"{area_col}{data_row_start}:{area_col}{data_row_end}"
        dept_range = f"{dept_col}{data_row_start}:{dept_col}{data_row_end}"

        area_formula = '"' + ",".join(AREA_OPTIONS) + '"'
        dv_area = DataValidation(type="list", formula1=area_formula, allow_blank=True)
        ws.add_data_validation(dv_area)
        dv_area.add(area_range)

        if CBSA_DEPARTMENTS:
            validation_sheet_name = "Listas"
            if validation_sheet_name in wb.sheetnames:
                val_sheet = wb[validation_sheet_name]
                val_sheet.delete_rows(1, val_sheet.max_row)
            else:
                val_sheet = wb.create_sheet(validation_sheet_name)

            for idx, dept_value in enumerate(CBSA_DEPARTMENTS, start=1):
                val_sheet.cell(row=idx, column=1, value=dept_value)
            val_sheet.sheet_state = "hidden"

            if "CBSA_Departments" in wb.defined_names:
                del wb.defined_names["CBSA_Departments"]

            dept_range_ref = f"'{validation_sheet_name}'!$A$1:$A${len(CBSA_DEPARTMENTS)}"
            defined_name = DefinedName(name="CBSA_Departments", attr_text=dept_range_ref)
            if hasattr(wb.defined_names, "append"):
                wb.defined_names.append(defined_name)
            else:
                wb.defined_names.add(defined_name)

            dept_formula = (
                f'=IF(INDIRECT("${area_col}"&ROW())="{AREA_FONDOS}","-",'
                f'IF(INDIRECT("${area_col}"&ROW())="{AREA_CBSA}",CBSA_Departments,""))'
            )
            dv_dept = DataValidation(type="list", formula1=dept_formula, allow_blank=True)
            ws.add_data_validation(dv_dept)
            dv_dept.add(dept_range)
        else:
            log("Aviso: la lista de departamentos CBSA está vacía; no se creó la validación de datos para 'Departamento'.")

    log(f"Listo. Archivo generado: {out_path}")

if __name__ == "__main__":
    main()
