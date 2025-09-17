# Write the complete script to a file the user can download and run locally.
script_path = "/mnt/data/tickets_parser.py"
script_code = r'''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tickets_parser.py
-----------------
Extrae campos de PDFs de tickets (formato Service Desk similar al que enviaste) y
genera un Excel en español con columnas útiles + tiempos calculados.

Reglas clave:
- Título: primera línea "real" después de `Urgency:` (saltando el valor de urgencia).
  Si inmediatamente aparece `Related client number or user:` (o variantes), se salta esa
  línea y la siguiente (su valor). La próxima línea no vacía ni timestamp es el título.
  El texto se respeta tal cual (no se elimina "Re:").
- Última respuesta: último mensaje HUMANO detectado por timestamps en el hilo;
  si no se detecta, se usa la última línea "de contenido" del texto limpio.
- Tiempos:
  * "Tiempo parado desde la última respuesta" = ahora(América/Asunción por defecto) - última respuesta.
  * "Tiempo abierto (si sigue abierto)" = ahora - fecha de creación, SOLO si el estado NO es cerrado/resuelto/etc.
- Parsing de fechas robusto: intenta DD/MM (por defecto). Si queda en el futuro, prueba MM/DD.
  Soporta 12h (AM/PM) y 24h. Si aún queda futuro, se descarta (no se muestra "0m").

Uso:
  python tickets_parser.py --pdf_dir "/ruta/a/pdfs" --out "/ruta/salida/Tickets.xlsx"
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
from datetime import datetime
from typing import Dict, List, Optional
from zoneinfo import ZoneInfo
from glob import glob

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

# Timestamps tipo "25/12/2024, 09:15 AM" o "12/25/2024, 09:15" (24h o 12h con/ sin segundos)
STAMP_RE = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{2,4}),?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?", re.I)
# Para localizar hilos
TS_LINE_RE = STAMP_RE

HEADER_BLOCK_RE = re.compile(
    r"Ticket\s*#(?P<ticket>\d+).*?Status\s+(?P<Status>.+?)\s+Name\s+(?P<Name>.+?)\s+Priority\s+(?P<Priority>.+?)\s+Email\s+(?P<Email>.+?)\s+Department\s+(?P<Department>.+?)\s+(?:Phone\s+.*?\s+)?Create Date\s+(?P<CreateDate>.+?)\s+Source\s+(?P<Source>.+?)\s+Ticket Details",
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

    urg = find_line_idx(lines, lambda ln: re.match(r"^\s*Urgency\s*:\s*$", ln.strip(), re.I) or re.match(r"^\s*Urgency\s*:\s*.+$", ln.strip(), re.I), start, 200)
    if urg is None:
        return ""

    i = urg + 1
    if i >= len(lines): return ""
    i += 1  # saltar valor de urgencia

    if i < len(lines) and is_related_label(lines[i]):
        i += 1  # etiqueta
        if i < len(lines): i += 1  # valor

    while i < len(lines):
        s = lines[i].strip()
        if s and not TS_LINE_RE.search(s):
            s = re.sub(r'^[\'"“”]+|[\'"“”]+$', "", s).strip()
            s = re.sub(r"\s+", " ", s)
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
        start = m.start()
        line_start = full_text.rfind("\n", 0, start) + 1
        line_end   = full_text.find("\n", m.end())
        if line_end == -1:
            line_end = len(full_text)
        header_line = full_text[line_start:line_end]
        author = header_line[m.end() - line_start:].strip()

        # si no hay autor en la misma línea, probar siguiente línea
        if not author or len(author) < 2:
            nxt_start = line_end + 1
            nxt_end   = full_text.find("\n", nxt_start)
            if nxt_end == -1: nxt_end = len(full_text)
            nxt_line  = (full_text[nxt_start:nxt_end] or "").strip()
            if nxt_line and not TS_LINE_RE.search(nxt_line) and not LABELY.match(nxt_line):
                author = nxt_line

        next_ts = matches[i + 1].start() if (i + 1) < len(matches) else len(full_text)
        body = full_text[line_end:next_ts].strip()
        entries.append({"stamp": m.group(0), "author": author, "body": body})
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
        if ampm.upper() == "PM" and hour != 12:
            hour += 12
        if ampm.upper() == "AM" and hour == 12:
            hour = 0
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
        if ampm.upper() == "PM" and hour != 12:
            hour += 12
        if ampm.upper() == "AM" and hour == 12:
            hour = 0
    try:
        return datetime(year, mo, d, hour, minute, sec, tzinfo=tz)
    except Exception:
        return None

def parse_pdf_timestamp(s: str, now: datetime, tz: ZoneInfo) -> Optional[datetime]:
    if not isinstance(s, str) or not s.strip():
        return None
    # por defecto DD/MM
    dt = parse_timestamp_ddmm(s, tz)
    if dt and dt > now:
        alt = parse_timestamp_mmdd(s, tz)
        if alt and alt <= now:
            dt = alt
        else:
            return None  # evitar "0m" por fechas futuras
    return dt

def parse_header_date(s: str, now: datetime, tz: ZoneInfo) -> Optional[datetime]:
    return parse_pdf_timestamp(s, now, tz)

def human_delta(td_seconds: Optional[float]) -> str:
    if td_seconds is None:
        return ""
    secs = int(td_seconds)
    if secs <= 0:
        return ""
    mins, _ = divmod(secs, 60)
    hrs, m = divmod(mins, 60)
    days, h = divmod(hrs, 24)
    if days > 0:
        return f"{days}d {h}h"
    if hrs > 0:
        return f"{hrs}h {m}m"
    return f"{m}m"

def is_open_status(status: str) -> bool:
    if not isinstance(status, str):
        return True
    st = status.strip().lower()
    closed_words = ["closed","resolved","solved","cerrado","resuelto","completado","finalizado","finalizada"]
    return not any(w in st for w in closed_words)

def parse_header_block(cleaned: str, lines: List[str]) -> Dict[str,str]:
    # 1) Intento bloque
    m = HEADER_BLOCK_RE.search(cleaned)
    if m:
        gd = m.groupdict()
        return {
            "Ticket Number": gd.get("ticket","").strip(),
            "Status":        gd.get("Status","").strip(),
            "Name":          gd.get("Name","").strip(),
            "Priority":      gd.get("Priority","").strip(),
            "Department":    gd.get("Department","").replace(" Phone","").strip(),
            "Create Date":   gd.get("CreateDate","").strip(),
        }
    # 2) Fallback línea a línea
    fields = {"Ticket Number":"","Status":"","Name":"","Priority":"","Department":"","Create Date":""}
    # Ticket #
    m2 = re.search(r"Ticket\s*#\s*(\d+)", cleaned, re.I)
    if m2: fields["Ticket Number"] = m2.group(1).strip()

    def cap(tok: str) -> str:
        for ln in lines[:120]:
            s = ln.strip()
            if s.lower().startswith(tok.lower() + " "):
                return s[len(tok)+1:].strip()
        return ""

    st = cap("Status")
    if " Name " in st:
        s, n = st.split(" Name ", 1)
        fields["Status"] = s.strip()
        fields["Name"]   = n.strip()
    else:
        fields["Status"] = st.strip()
        nm = cap("Name")
        if nm: fields["Name"] = nm

    pr  = cap("Priority");   fields["Priority"] = pr.split(" Email ",1)[0].strip() if pr else pr
    dep = cap("Department"); fields["Department"] = dep.replace(" Phone","").strip()
    cd  = cap("Create Date");fields["Create Date"] = cd.split(" Source ",1)[0].strip() if cd else cd

    return fields

def parse_pdf(pdf_path: str, tz: ZoneInfo, dump_dir: Optional[str], now: datetime) -> Dict[str,str]:
    raw = extract_text(pdf_path)
    if not raw or len(raw) < 80:
        return {"N° Ticket":"", "Título del ticket":"", "Estado":"", "Prioridad":"", "Departamento":"", "Fecha de creación":"", "Última respuesta por":"", "Última respuesta el":"", "Error":"no_text_extracted", "_src": os.path.basename(pdf_path)}
    cleaned = clean_text(raw)
    lines = cleaned.splitlines()

    if dump_dir:
        os.makedirs(dump_dir, exist_ok=True)
        base = os.path.basename(pdf_path)
        with open(os.path.join(dump_dir, base + ".raw.txt"), "w", encoding="utf-8") as f:
            f.write(raw)
        with open(os.path.join(dump_dir, base + ".cleaned.txt"), "w", encoding="utf-8") as f:
            for i, ln in enumerate(lines, start=1):
                f.write(f"{i:04d} | {ln}\n")

    hdr = parse_header_block(cleaned, lines)
    title = extract_title_after_urgency(cleaned)

    # Hilo → última respuesta humana
    entries = extract_thread_entries(cleaned)
    last_by, last_at = "", ""
    for e in entries:
        if e["author"] and not is_auto(e["author"], e["body"]):
            last_by = e["author"]
            last_at = e["stamp"]
    if not last_by:
        last_by = fallback_last_author_from_tail(cleaned)
    if not last_at:
        last_stamp = None
        for m in TS_LINE_RE.finditer(cleaned):
            last_stamp = m.group(0)
        if last_stamp:
            last_at = last_stamp

    return {
        "N° Ticket": hdr.get("Ticket Number",""),
        "Título del ticket": title,
        "Estado": hdr.get("Status",""),
        "Prioridad": hdr.get("Priority",""),
        "Departamento": hdr.get("Department",""),
        "Fecha de creación": hdr.get("Create Date",""),
        "Última respuesta por": last_by,
        "Última respuesta el": last_at,
        "Error": "",
        "_src": os.path.basename(pdf_path)
    }

def main():
    ap = argparse.ArgumentParser(description="PDF tickets -> Excel (ES) con título post-Urgency, última respuesta y tiempos.")
    ap.add_argument("--pdf_dir", required=True, help="Carpeta con PDFs")
    ap.add_argument("--out", required=True, help="Ruta de salida del Excel")
    ap.add_argument("--pattern", default="*.pdf", help="Patrón de archivos (glob). Ej: 'Ticket-*.pdf' o '*.pdf'")
    ap.add_argument("--timezone", default="America/Asuncion", help="Zona horaria (por defecto America/Asuncion)")
    ap.add_argument("--dedupe", action="store_true", help="Desduplicar por N° Ticket, conservando la fila más completa")
    ap.add_argument("--dump_text", action="store_true", help="Guardar textos crudo/limpio en debug_texts/")
    args = ap.parse_args()

    tz = ZoneInfo(args.timezone)
    now = datetime.now(tz)

    # Recolectar PDFs
    pattern = os.path.join(args.pdf_dir, args.pattern)
    files = sorted(glob(pattern))
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files:
        log("No se encontraron PDFs con el patrón indicado.")
        sys.exit(1)

    dump_dir = os.path.join("debug_texts") if args.dump_text else None

    log(f"Encontrados {len(files)} PDF(s). Procesando…")
    rows = []
    for idx, path in enumerate(files, start=1):
        log(f"[{idx}/{len(files)}] {os.path.basename(path)}")
        try:
            row = parse_pdf(path, tz, dump_dir, now)
            rows.append(row)
        except KeyboardInterrupt:
            log("Interrumpido por el usuario (Ctrl+C). Guardando progreso parcial…")
            break
        except Exception as e:
            log(f"   ! Error en {os.path.basename(path)}: {e}")
            rows.append({"N° Ticket":"", "Título del ticket":"", "Estado":"", "Prioridad":"", "Departamento":"",
                         "Fecha de creación":"", "Última respuesta por":"", "Última respuesta el":"",
                         "Error": f"{os.path.basename(path)}: {e}", "_src": os.path.basename(path)})

    if not rows:
        log("No hay filas para exportar.")
        sys.exit(1)

    df = pd.DataFrame(rows)

    # Tiempos calculados
    def parse_pdf_timestamp_for_row(s: str) -> Optional[datetime]:
        return parse_pdf_timestamp(s, now, tz)

    def parse_header_date_for_row(s: str) -> Optional[datetime]:
        return parse_header_date(s, now, tz)

    def human_delta_from(dt: Optional[datetime]) -> str:
        if not dt:
            return ""
        return human_delta((now - dt).total_seconds())

    df["Tiempo parado desde la última respuesta"] = df["Última respuesta el"].apply(lambda s: human_delta_from(parse_pdf_timestamp_for_row(s)))
    def compute_open_age(row):
        st = row.get("Estado","")
        cd = parse_header_date_for_row(row.get("Fecha de creación",""))
        if cd and is_open_status(st):
            return human_delta((now - cd).total_seconds())
        return ""
    df["Tiempo abierto (si sigue abierto)"] = df.apply(compute_open_age, axis=1)

    # Desduplicar (opcional)
    if args.dedupe and "N° Ticket" in df.columns:
        def score_row(row):
            return sum(1 for v in row if (isinstance(v, str) and v.strip()) or (pd.notna(v) and v != ""))
        df["_score"] = df.apply(score_row, axis=1)
        df = df.sort_values(["_score"], ascending=[False]).drop_duplicates(subset=["N° Ticket"], keep="first")
        df = df.drop(columns=["_score"], errors="ignore")

    # Orden de columnas
    final_cols = [
        "N° Ticket","Título del ticket","Estado","Prioridad","Departamento",
        "Fecha de creación","Última respuesta por","Última respuesta el",
        "Tiempo parado desde la última respuesta","Tiempo abierto (si sigue abierto)"
    ]
    for c in final_cols:
        if c not in df.columns: df[c] = ""
    df = df[final_cols + ["_src","Error"]]

    # Guardar
    out_path = args.out
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    df.to_excel(out_path, index=False)
    log(f"Listo. Archivo generado: {out_path}")

if __name__ == "__main__":
    main()
'''
with open(script_path, "w", encoding="utf-8") as f:
    f.write(script_code)

script_path
