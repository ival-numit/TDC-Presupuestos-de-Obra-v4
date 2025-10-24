# parser_presupuesto.py (pandas-free, writes Excel via XlsxWriter)
import os, re, io
import pdfplumber
from datetime import datetime
import xlsxwriter

# ========= Utilidades =========
MESES = {"enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
         "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12}
def parse_num(s):
    if not s: return None
    s = s.replace(",", "").replace("$","").strip()
    try: return float(s)
    except: return None
def to_iso_date(text):
    m = re.search(r"(\d{1,2})\s+de\s+([a-záéíóú]+)\s+de\s+(\d{4})", (text or "").lower())
    if not m: return None
    d, mes, y = int(m.group(1)), m.group(2), int(m.group(3))
    return datetime(y, MESES.get(mes,1), d).date().isoformat()

# ========= Patrones =========
UNIT_TOKEN = r"[A-Za-zÁÉÍÓÚáéíóúñÑº°/\.\-\u00B2\u00B30-9]{1,16}"
RE_SECTION_CORE = re.compile(r"^(\d{1,2}(?:\.\d{1,2})?)\s+(.+?)\s+([\d\.,]+)$", re.UNICODE)
UNIT_LIKE = re.compile(r"\b(m2|m3|cm|mm|ml|kg|kg/m2|kg/m3|pza\.?|pz\.?|pieza|lote|vj|lts?|litros?)\b", re.IGNORECASE)
DIGIT_COLON_DIGIT = re.compile(r"\d\s*:\s*\d")

def try_match_section(line: str):
    line = line.strip()
    if DIGIT_COLON_DIGIT.search(line):
        return None
    m = RE_SECTION_CORE.match(line)
    if not m: return None
    code, name, total = m.groups()
    name = name.strip()
    try:
        if "." in code:
            a, b = map(int, code.split("."))
            if not (1 <= a <= 99 and 1 <= b <= 99): return None
        else:
            if not (1 <= int(code) <= 99): return None
    except:
        return None
    if len(name) < 3 or UNIT_LIKE.search(name):
        return None
    return code, name, total

def is_header_or_total(line: str) -> bool:
    ln = (line or "").lower()
    if ln.startswith("clave descripción") or "osroca" in ln: return True
    if re.match(r"^(subtotal|total|iva|notas|ppro|ppto)", ln, re.IGNORECASE): return True
    if ln.startswith("página ") or re.match(r"^\d+/\d+$", ln): return True
    return False

# --- CLAVE ---
KEY_TDC = r"TDC-(?:[A-Z]{1,3})(?:-[A-Z]{1,3})?-?"
KEY_GENERIC_CORE = r"(?=.*[A-Z])(?:[A-Z0-9]{1,6}(?:[.\-/][A-Z0-9]{1,6})*)"
KEY_GENERIC = rf"(?:{KEY_GENERIC_CORE})(?:\s+{KEY_GENERIC_CORE})?"
KEY_TOKEN = rf"(?:{KEY_TDC}|{KEY_GENERIC})"

RE_PARTIDA_1LINE = re.compile(
    rf"^({KEY_TOKEN})\s+(.*?)\s+({UNIT_TOKEN})\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)(?:\s+.*)?$",
    re.UNICODE
)
RE_CODE_ONLY   = re.compile(rf"^({KEY_TOKEN})\s+(.*)$", re.UNICODE)
RE_VALUES_LINE = re.compile(rf"^({UNIT_TOKEN})\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)$", re.UNICODE)
RE_THREE_NUMS_AT_END = re.compile(r"([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s*$")

def parse_pdf(pdf_path):
    rows = []
    titulo = fecha_iso = None
    sec_code = sec_name = None
    sub_code = sub_name = None
    pending_code = None
    pending_desc_parts = []
    append_to_last = False
    unmatched_codes = []

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=1.5, y_tolerance=2.0) or ""
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

            if i <= 2:
                for ln in lines:
                    if not titulo and ln.lower().startswith("presupuesto"): titulo = ln
                    if not fecha_iso: fecha_iso = to_iso_date(ln) or fecha_iso

            for ln in lines:
                if is_header_or_total(ln):
                    pending_code = None; pending_desc_parts = []; append_to_last = False
                    continue

                sec = try_match_section(ln)
                if sec:
                    code, name, _ = sec
                    if "." in code:
                        sub_code, sub_name = code, name
                    else:
                        sec_code, sec_name = code, name
                        sub_code = sub_name = None
                    pending_code = None; pending_desc_parts = []; append_to_last = False
                    continue

                m1 = RE_PARTIDA_1LINE.match(ln)
                if m1:
                    clave, desc, unidad, cant, pu, tot = m1.groups()[:6]
                    rows.append({
                        "seccion": sec_code, "seccion_nombre": sec_name,
                        "subseccion": sub_code, "subseccion_nombre": sub_name,
                        "clave": clave, "descripcion": (desc or "").strip(),
                        "unidad": unidad.strip(),
                        "cantidad": parse_num(cant),
                        "precio_unitario": parse_num(pu),
                        "total": parse_num(tot),
                        "titulo": titulo, "fecha": fecha_iso,
                        "archivo": os.path.basename(pdf_path)
                    })
                    append_to_last = True
                    pending_code = None; pending_desc_parts = []
                    continue

                mcode = RE_CODE_ONLY.match(ln)
                if mcode:
                    pending_code = mcode.group(1)
                    pending_desc_parts = [mcode.group(2).strip()] if mcode.group(2) else []
                    append_to_last = False
                    continue

                mvals = RE_VALUES_LINE.match(ln)
                if mvals and pending_code:
                    unidad, cant, pu, tot = mvals.groups()
                    desc_full = " ".join(pending_desc_parts).strip() if pending_desc_parts else None
                    rows.append({
                        "seccion": sec_code, "seccion_nombre": sec_name,
                        "subseccion": sub_code, "subseccion_nombre": sub_name,
                        "clave": pending_code, "descripcion": desc_full,
                        "unidad": unidad.strip(),
                        "cantidad": parse_num(cant),
                        "precio_unitario": parse_num(pu),
                        "total": parse_num(tot),
                        "titulo": titulo, "fecha": fecha_iso,
                        "archivo": os.path.basename(pdf_path)
                    })
                    pending_code = None; pending_desc_parts = []; append_to_last = True
                    continue

                looks_like_code = re.match(rf"^{KEY_TOKEN}\b", ln) is not None
                is_real_one_line = RE_PARTIDA_1LINE.match(ln) is not None
                is_real_code_only = RE_CODE_ONLY.match(ln) is not None
                ends_with_three_nums = RE_THREE_NUMS_AT_END.search(ln) is not None

                if append_to_last and rows and not try_match_section(ln):
                    if looks_like_code and not (is_real_one_line or is_real_code_only or ends_with_three_nums):
                        rows[-1]["descripcion"] = (rows[-1]["descripcion"] + " " + ln).strip()
                        continue
                    if not looks_like_code and not is_real_one_line:
                        rows[-1]["descripcion"] = (rows[-1]["descripcion"] + " " + ln).strip()
                        continue

                if looks_like_code and not (is_real_one_line or is_real_code_only):
                    unmatched_codes.append(ln)
                    append_to_last = False
                    pending_code = None; pending_desc_parts = []
                    continue

    return rows, unmatched_codes

def build_xlsx_result(rows, out_xlsx_name="presupuesto_bd.xlsx"):
    headers = ['seccion','seccion_nombre','subseccion','subseccion_nombre','clave',
               'descripcion','unidad','cantidad','precio_unitario','total','titulo','fecha','archivo']
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('BD')
    header_fmt = wb.add_format({'bold': True})
    for col, h in enumerate(headers):
        ws.write(0, col, h, header_fmt)
    for r, row in enumerate(rows, start=1):
        ws.write(r, 0, row.get('seccion'))
        ws.write(r, 1, row.get('seccion_nombre'))
        ws.write(r, 2, row.get('subseccion'))
        ws.write(r, 3, row.get('subseccion_nombre'))
        ws.write(r, 4, row.get('clave'))
        ws.write(r, 5, row.get('descripcion'))
        ws.write(r, 6, row.get('unidad'))
        def num(v):
            try:
                return float(v) if v is not None else None
            except:
                return None
        ws.write_number(r, 7, num(row.get('cantidad')) or 0)
        ws.write_number(r, 8, num(row.get('precio_unitario')) or 0)
        ws.write_number(r, 9, num(row.get('total')) or 0)
        ws.write(r, 10, row.get('titulo'))
        ws.write(r, 11, row.get('fecha'))
        ws.write(r, 12, row.get('archivo'))
    ws.autofilter(0, 0, max(len(rows),1), len(headers)-1)
    ws.freeze_panes(1, 0)
    for c in range(len(headers)):
        ws.set_column(c, c, 18)
    wb.close()
    output.seek(0)
    return output
