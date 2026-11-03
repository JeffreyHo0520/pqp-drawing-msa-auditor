# /// script
# dependencies = [
#   "openpyxl",
#   "pymupdf",
#   "rapidocr-onnxruntime",
#   "pillow"
# ]
# ///
import openpyxl
import os
import sys
import json
import re
import argparse
import io
import zipfile
import xml.etree.ElementTree as ET
from PIL import Image

try:
    import openpyxl.drawing.text as dt
    if hasattr(dt, 'Font') and hasattr(dt.Font, 'pitchFamily'):
        dt.Font.pitchFamily.max = 255
except Exception:
    pass

def reconfigure_stdout():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')

def read_file_bytes_safe(path):
    try:
        with open(path, 'rb') as f:
            return f.read()
    except (PermissionError, IOError):
        if os.name == 'nt':
            try:
                import ctypes
                from ctypes import wintypes
                kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
                handle = kernel32.CreateFileW(
                    os.path.abspath(path),
                    0x80000000, # GENERIC_READ
                    1 | 2 | 4,  # FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE
                    None,
                    3,          # OPEN_EXISTING
                    0x80,       # FILE_ATTRIBUTE_NORMAL
                    None
                )
                if handle != -1:
                    size = kernel32.GetFileSize(handle, None)
                    buf = ctypes.create_string_buffer(size)
                    bytes_read = wintypes.DWORD()
                    kernel32.ReadFile(handle, buf, size, ctypes.byref(bytes_read), None)
                    kernel32.CloseHandle(handle)
                    return buf.raw[:bytes_read.value]
            except Exception:
                pass
        raise

def load_workbook_safe(path, **kwargs):
    data = read_file_bytes_safe(path)
    return openpyxl.load_workbook(io.BytesIO(data), **kwargs)

def get_fmt_decimals(fmt):
    """
    Extracts the number of decimal places defined in an Excel number format string
    (e.g., '0.000' -> 3, '0.00_ ' -> 2, '0.0' -> 1, 'General' -> None).
    """
    if not fmt or fmt == 'General':
        return None
    clean = str(fmt).replace('_', '').replace(' ', '')
    m = re.search(r'\.(0+)', clean)
    if m:
        return len(m.group(1))
    return 0

# Helper: Parse fraction

def parse_float_val(cell, metric_type=None):
    """
    Parses a cell value to float.
    Handles '8%', '100%', fractions like 0.08, 1.0, and flags error cells (#DIV/0!, #N/A).
    """
    if cell is None:
        raise ValueError("Cell is missing")
        
    if hasattr(cell, 'value'):
        val = cell.value
        num_format = getattr(cell, 'number_format', 'General') or 'General'
    else:
        val = cell
        num_format = 'General'
        
    if val is None:
        raise ValueError("Cell is empty")
        
    val_str = str(val).strip()
    if val_str == "" or val_str.startswith("#"):
        raise ValueError(f"Cell error or blank value: '{val_str}'")
        
    has_percent = '%' in val_str or '%' in str(num_format)
    val_clean = val_str.replace('%', '').strip()
    
    try:
        val_float = float(val_clean)
    except ValueError:
        raise ValueError(f"Value '{val_str}' is not a valid number")
        
    if has_percent:
        if metric_type in ["% Contribution", "% Study Var", "% Tolerance", "SCREEN % EFFECTIVE SCORE", "ATTRIBUTE (4)"]:
            if val_float <= 1.0:
                val_float = val_float * 100.0
                
    return val_float

def resolve_formula_value(wb, sheet, val_str):
    """
    Resolves basic formula references across sheets (e.g., =TitlePg!D7, =PQS!C12, =IF(ISBLANK(...))).
    """
    if not isinstance(val_str, str) or not val_str.startswith("="):
        return val_str
        
    val_upper = val_str.upper().strip()
    
    m_if = re.search(r"ISBLANK\(([^!]+)!([A-Z0-9]+)\)", val_upper)
    if m_if:
        sheet_name_raw = m_if.group(1).strip().strip("'")
        cell_coord = m_if.group(2).strip()
        sheet_name = re.sub(r"^\[\d+\]", "", sheet_name_raw).strip()
        
        target_sheet = None
        if sheet_name in wb.sheetnames:
            target_sheet = wb[sheet_name]
        else:
            for name in wb.sheetnames:
                if name.lower().strip() == sheet_name.lower():
                    target_sheet = wb[name]
                    break
        if target_sheet:
            target_val = target_sheet[cell_coord].value
            if target_val is not None:
                resolved = resolve_formula_value(wb, target_sheet, str(target_val))
                if resolved == "" or str(resolved).strip() == "":
                    return " "
                return resolved

    m_ref = re.search(r"^=([^!]+)!([A-Z0-9]+)$", val_upper)
    if m_ref:
        sheet_name_raw = m_ref.group(1).strip().strip("'")
        cell_coord = m_ref.group(2).strip()
        sheet_name = re.sub(r"^\[\d+\]", "", sheet_name_raw).strip()
        
        target_sheet = None
        if sheet_name in wb.sheetnames:
            target_sheet = wb[sheet_name]
        else:
            for name in wb.sheetnames:
                if name.lower().strip() == sheet_name.lower():
                    target_sheet = wb[name]
                    break
        if target_sheet:
            target_val = target_sheet[cell_coord].value
            if target_val is not None:
                return resolve_formula_value(wb, target_sheet, str(target_val))
                
    m_single = re.search(r"^=([A-Z0-9]+)$", val_upper)
    if m_single:
        cell_coord = m_single.group(1).strip()
        target_val = sheet[cell_coord].value
        if target_val is not None:
            return resolve_formula_value(wb, sheet, str(target_val))
            
    return val_str

def parse_fraction(s):
    s = s.strip()
    if '-' in s:
        parts = s.split('-')
        whole = float(parts[0])
        frac_parts = parts[1].split('/')
        frac = float(frac_parts[0]) / float(frac_parts[1])
        return whole + frac
    elif '/' in s:
        frac_parts = s.split('/')
        return float(frac_parts[0]) / float(frac_parts[1])
    else:
        return float(s)

def parse_nom_tol(s, expected_val=None):
    # Clean up s
    s = s.strip().replace(',', '.').strip('()[]{}')
    s = re.sub(r'^[RrQq\-\+]+', '', s) # remove Q, R, -, +
    s = re.sub(r'^\d+-(?!\d+/\d+)', '', s)
    if s.startswith('[') and s.endswith(']'):
        s = s[1:-1].strip()
    s = re.sub(r'\[\s*\]$', '', s)
    s = re.sub(r'\[\]$', '', s)
    
    # Try parsing as a fraction
    frac_match = re.search(r'(\d+-\d+/\d+|\d+/\d+)', s)
    if frac_match:
        try:
            val = parse_fraction(frac_match.group(1))
            return val, None
        except Exception:
            pass
            
    # Check if it can be directly parsed as a float and matches expected_val close enough
    try:
        val = float(s)
        if expected_val is not None:
            if abs(val - expected_val) < 0.05 * expected_val or abs(val - expected_val) < 0.01:
                return val, None
        else:
            return val, None
    except ValueError:
        pass

    # 1. Look for explicit ± or +/-
    for sep in ['±', '+/-', '±', '#']:
        if sep in s:
            parts = s.split(sep)
            try:
                return float(parts[0].strip()), float(parts[1].strip())
            except ValueError:
                pass

    # 1.1 Space-separated (e.g. '.75 .02' or '1.850 .020' or '19.1 0.5')
    if ' ' in s:
        parts = [p.strip() for p in s.split() if p.strip()]
        if len(parts) == 2:
            try:
                p0 = float(parts[0])
                p1 = float(parts[1])
                return p0, p1
            except ValueError:
                pass
                
    # 2. If there is no explicit separator, but expected_val is given:
    if expected_val is not None:
        clean_s = re.sub(r'[^\d\.]', '', s)
        best_nom = None
        best_tol = None
        min_diff = 999.0
        
        dot_parts = clean_s.split('.')
        if len(dot_parts) > 2:
            try:
                nom_candidate = float(dot_parts[0] + "." + dot_parts[1])
                tol_candidate = float("0." + dot_parts[2])
                diff = abs(nom_candidate - expected_val)
                if diff < min_diff:
                    min_diff = diff
                    best_nom = nom_candidate
                    best_tol = tol_candidate
            except ValueError:
                pass
        else:
            for split_idx in range(1, len(clean_s)):
                part1 = clean_s[:split_idx]
                part2 = clean_s[split_idx:]
                try:
                    nom_cand = float(part1)
                    tol_cand = float(part2)
                    if tol_cand >= 1.0:
                        tol_cand = tol_cand / (10 ** len(part2))
                    diff = abs(nom_cand - expected_val)
                    if diff < min_diff and diff < 0.2 * expected_val:
                        min_diff = diff
                        best_nom = nom_cand
                        best_tol = tol_cand
                except ValueError:
                    pass
        if best_nom is not None and min_diff < 0.15:
            return best_nom, best_tol

    # 3. Simple float conversion fallback (supporting leading dot like .43 or .032)
    try:
        float_match = re.search(r'(\d*\.\d+|\d+)', s)
        if float_match and float_match.group(1):
            return float(float_match.group(1)), None
    except ValueError:
        pass
        
    return None, None


class DrawingParser:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.filename = os.path.basename(pdf_path)
        self.pages_text = []
        self.load_pdf()

    def load_pdf(self):
        try:
            try:
                import pymupdf
                doc = pymupdf.open(self.pdf_path)
                for page in doc:
                    self.pages_text.append(page.get_text() or "")
            except ImportError:
                import pypdf
                reader = pypdf.PdfReader(self.pdf_path)
                for page in reader.pages:
                    self.pages_text.append(page.extract_text() or "")
                doc = None
            
            # Check if we should also run OCR
            combined_text = "".join(self.pages_text).strip()
            dwg_no = self.extract_metadata_from_text(combined_text)
            
            # If no text was extracted, or drawing number is not found, attempt OCR
            if not combined_text or len(combined_text) < 20 or not dwg_no:
                print(f"No text or drawing number extracted via pymupdf from {self.filename}. Attempting OCR...", flush=True)
                ocr_texts = []
                from rapidocr_onnxruntime import RapidOCR
                engine = RapidOCR()
                
                for i, page in enumerate(doc):
                    pix = page.get_pixmap(dpi=300)
                    import tempfile
                    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                        img_path = tmp.name
                    try:
                        pix.save(img_path)
                        results, elapse = engine(img_path)
                        page_texts = []
                        if results:
                            for res in results:
                                box, text, score = res
                                page_texts.append(text)
                        ocr_texts.append(" ".join(page_texts))
                    finally:
                        if os.path.exists(img_path):
                            os.remove(img_path)
                
                if ocr_texts:
                    self.pages_text.append(" ".join(ocr_texts))
                    print(f"OCR extracted {len(' '.join(ocr_texts))} chars from {self.filename}.", flush=True)
        except Exception as e:
            print(f"Error loading PDF {self.filename}: {e}", file=sys.stderr)

    def extract_metadata_from_text(self, text):
        dwg_no = None
        invalid_tokens = {'ITER', 'REV', 'TITLE', 'SIZE', 'SCALE', 'SHEET', 'ZONE', 'ITEM', 'DESCRIPTION', 'A', 'B', 'C', 'D', 'E'}
        
        # 1. Match title block with DWG NO / DRAWING NO followed by valid drawing number
        tb_match = re.search(r'DO NOT SCALE\s*\n?\s*DRAWING\s*\n?\s*([0-9A-Z-]+)', text, re.IGNORECASE)
        if tb_match and tb_match.group(1).strip().upper() not in invalid_tokens:
            dwg_no = tb_match.group(1).strip()
            
        # 2. Check filename for standard ASSA ABLOY patterns
        base_name = os.path.splitext(self.filename)[0]
        if not dwg_no:
            fn_pat = re.search(r'\b(\d{3}-\d{5}-\d{3})\b', base_name) or re.search(r'\b(\d{5}-\d{3})\b', base_name)
            if fn_pat:
                dwg_no = fn_pat.group(1)

        if not dwg_no:
            dwg_match = re.search(r'\b(KPZ\d{5,7})\b', text, re.IGNORECASE)
            if dwg_match:
                dwg_no = dwg_match.group(1).upper()
            else:
                dwg_match_dash = re.search(r'\b(\d{3}-\d{5}-\d{3})\b', text)
                if dwg_match_dash:
                    dwg_no = dwg_match_dash.group(1)
                else:
                    dwg_match_dash2 = re.search(r'\b(\d{5}-\d{3})\b', text)
                    if dwg_match_dash2:
                        dwg_no = dwg_match_dash2.group(1)
                    else:
                        dwg_match2 = re.search(r'\b([A-Z]{3,4}\d{5,7})\b', text)
                        if dwg_match2:
                            dwg_no = dwg_match2.group(1).upper()
                        else:
                            dwg_match3 = re.search(r'\b(5\d{6,7})\b', text)
                            if dwg_match3:
                                dwg_no = dwg_match3.group(1)
                            else:
                                dwg_match_raw = re.search(r'DRAWING\s+NO\.?\s*([A-Z0-9-]+)', text, re.IGNORECASE)
                                if dwg_match_raw and dwg_match_raw.group(1).strip().upper() not in invalid_tokens:
                                    dwg_no = dwg_match_raw.group(1).strip()
                                else:
                                    dwg_match_raw2 = re.search(r'DWG\s+NO\.?\s*([A-Z0-9-]+)', text, re.IGNORECASE)
                                    if dwg_match_raw2 and dwg_match_raw2.group(1).strip().upper() not in invalid_tokens:
                                        dwg_no = dwg_match_raw2.group(1).strip()
                                        
        if not dwg_no or dwg_no.upper() in invalid_tokens:
            fn_match = re.search(r'\b(KPZ\d{5,7})\b', base_name, re.IGNORECASE)
            if fn_match:
                dwg_no = fn_match.group(1).upper()
            else:
                fn_dash = re.search(r'\b(\d{3}-\d{5}-\d{3})\b', base_name) or re.search(r'\b(\d{5}-\d{3})\b', base_name)
                if fn_dash:
                    dwg_no = fn_dash.group(1)
                else:
                    base_name_clean = re.sub(r'[_\s-]+\d+$', '', base_name)
                    file_nums = "".join(filter(str.isdigit, base_name_clean))
                    if len(file_nums) >= 5:
                        dwg_no = file_nums
                    else:
                        if len(base_name_clean) >= 5:
                            dwg_no = base_name_clean.upper()
                        else:
                            dwg_no = base_name.upper()
        return dwg_no

    def extract_metadata(self):
        text = " ".join(self.pages_text)
        dwg_no = self.extract_metadata_from_text(text)
        
        # 2. Revision
        revision = None
        
        # 2.0 Check filename for explicit revision (e.g. '50992550 3.pdf' -> 3, '5052610 6.pdf' -> 6)
        fn_rev_match = re.search(r'[_\s]+(\d+)(?:\s*\(\d+\))?\.pdf$', self.filename, re.IGNORECASE)
        if fn_rev_match:
            revision = fn_rev_match.group(1)
            
        # 2.05 Check ASSA ABLOY title block: "Production \n <REV> \n <ITER>"
        if not revision:
            prod_state_match = re.search(r'Production\s+(\d+|[A-Z])\s+\d+', text, re.IGNORECASE)
            if prod_state_match:
                revision = prod_state_match.group(1).upper()

        # 2.1 Check for production release revision (e.g. 1 PRODUCTION RELEASE / 2 PRODUCTION RELEASED)
        if not revision:
            prod_revs = re.findall(r'\b(\d+|[A-Z])\s+PRODUCTION\s+RELEASE[D]?', text, re.IGNORECASE)
            if prod_revs:
                num_revs = [int(x) for x in prod_revs if x.isdigit()]
                if num_revs:
                    revision = str(max(num_revs))
                else:
                    revision = prod_revs[-1].upper()
            
        # 2.2 Check for title block REV / Sheet rev
        if not revision:
            tb_rev = re.search(r'\bREV\s*\n?\s*([A-Z\d]+)\b', text, re.IGNORECASE)
            if tb_rev and tb_rev.group(1).upper() not in ["NO", "NUM", "ITER", "DESCRIPTION", "RELEASE"]:
                cand = tb_rev.group(1).upper()
                if cand in ["A", "B"] and re.search(r'\b' + cand + r'\s+B\s+C\s+D\b', text):
                    pass # ignore drawing frame zone grid coordinates
                else:
                    revision = cand
                
        if not revision:
            rev_matches = re.findall(r'\bREV\s*[:\-\s]?\s*([A-Z\d-]+)\b', text, re.IGNORECASE)
            for r_val in rev_matches:
                if r_val.upper() not in ["NO", "NUM", "ITER", "DESCRIPTION", "RELEASE"]:
                    if r_val.upper() in ["A", "B"] and re.search(r'\b' + r_val.upper() + r'\s+B\s+C\s+D\b', text):
                        continue
                    revision = r_val.upper()
                    if r_val.isdigit():
                        break
                
        if not revision:
            for line in self.pages_text:
                for line_part in line.split('\n'):
                    line_part = line_part.strip()
                    if re.match(r'^[A-Z]-\d+$', line_part): # e.g. B-1
                        revision = line_part.split('-')[0] # extract B
                        break
                    elif line_part in ["A", "B", "C", "D", "E"]:
                        revision = line_part
                        break
                if revision:
                    break

        if not revision:
            rev_match = re.search(r'REV\s*(\d+)', text, re.IGNORECASE)
            if rev_match:
                revision = rev_match.group(1)
            else:
                revision = "1"
                
        # 3. Weight
        weight = None
        weight_match = re.search(r'WEIGHT:\s*([\d\.]+)\s*(G|LBS|G\.)', text, re.IGNORECASE)
        if weight_match:
            weight = f"{weight_match.group(1)} {weight_match.group(2)}"

        # 4. Title / Description
        title = None
        m_title = re.search(r'([A-Z0-9\s,-]+)\n\s*[A-E]\s*\n\s*DO NOT SCALE\s*\n\s*DRAWING\s*\n\s*([0-9-]+)', text)
        if m_title:
            cand = m_title.group(1).splitlines()[-1].strip()
            if len(cand) >= 4 and not any(k in cand.upper() for k in ["SIZE", "DWG NO", "SCALE", "SHEET", "DRAFTER", "CHECKED"]):
                title = cand
        if not title:
            m_dwg = re.search(r'\bTITLE\b[\s\S]{1,120}?(INSERT[^\r\n]+|DIVIDER[^\r\n]+|PCBA[^\r\n]+|[A-Z0-9\s,-]{5,40})(?:\s+SIZE|\s+C|\s+D|\s+DO NOT SCALE|\s+DWG|\s+REV)', text)
            if m_dwg:
                cand = m_dwg.group(1).strip()
                if not any(k in cand.upper() for k in ["SIZE", "DWG NO", "SCALE", "SHEET", "DRAFTER", "CHECKED"]):
                    title = cand

        return {
            "drawing_number": dwg_no,
            "revision": revision,
            "weight": weight,
            "title": title
        }

    def extract_torques(self):
        text = " ".join(self.pages_text)
        torques = []
        
        matches = re.finditer(r'TORQUE\s+(?:IS|TO)?\s*([\d\.]+)\s*(?:±|\+/-)?\s*([\d\.]*)\s*(N-M|IN-LBS|NM|IN-LB)', text, re.IGNORECASE)
        for m in matches:
            val = float(m.group(1))
            tol = float(m.group(2)) if m.group(2) else 0.1
            unit = m.group(3).upper()
            torques.append({
                "nominal": val,
                "tolerance": tol,
                "unit": unit,
                "raw_text": m.group(0)
            })
            
        return torques

    def extract_sc_notes(self):
        """
        Extracts Significant Characteristic (SC) notes flagged with diamond shapes (◇ / ◊ / ◆)
        or explicit SC callouts from engineering drawing PDFs.
        """
        try:
            import pymupdf
            doc = pymupdf.open(self.pdf_path)
            sc_notes = []
            for page_idx, page in enumerate(doc):
                paths = page.get_drawings()
                diamond_rects = []
                for p in paths:
                    r = p["rect"]
                    w = r.width
                    h = r.height
                    # A diamond path has 4 segments or nearly equal width and height
                    if 6 <= w <= 22 and 6 <= h <= 22 and abs(w - h) <= 4:
                        diamond_rects.append(r)
                
                page_dict = page.get_text("dict")
                lines_with_text = []
                for block in page_dict.get("blocks", []):
                    if "lines" in block:
                        for line in block["lines"]:
                            text = "".join(s["text"] for s in line["spans"]).strip()
                            bbox = line["bbox"]
                            if text:
                                lines_with_text.append((bbox, text))
                
                for bbox, text in lines_with_text:
                    m_explicit = re.match(r'^(?:◇|◊|◆|\s*\[?\s*SC\s*\]?\s*)(\d+)\.\s*(.*)', text, re.IGNORECASE)
                    m_general = re.match(r'^(\d+)\.\s+(.*)', text)
                    
                    if m_explicit:
                        note_num = int(m_explicit.group(1))
                        note_content = m_explicit.group(2).strip()
                        is_sc_diamond = True
                    elif m_general:
                        note_num = int(m_general.group(1))
                        note_content = m_general.group(2).strip()
                        # Must look like an actual note sentence, not decimal numbers
                        if len(note_content.split()) < 2 or note_content.replace('.', '').isdigit():
                            continue
                        is_sc_diamond = False
                        for dr in diamond_rects:
                            # Diamond immediately preceding the note number
                            if (bbox[0] - 25) <= dr.x1 <= (bbox[0] + 5) and abs(dr.y0 - bbox[1]) < 6:
                                is_sc_diamond = True
                                break
                    else:
                        continue
                        
                    if is_sc_diamond:
                        # Skip legend / explanatory notes defining what SC/CTQ means
                        upper_cont = note_content.upper()
                        if "SIGNIFICANT CHARACTERISTIC" in upper_cont or "CPK REQUIRED" in upper_cont or "CTQ PER" in upper_cont:
                            continue
                        sc_notes.append({
                            "note_num": note_num,
                            "content": note_content,
                            "raw_line": text,
                            "page": page_idx + 1
                        })
            return sc_notes
        except Exception:
            return []

    def extract_sc_features(self):
        """
        Extracts Significant Characteristics (SC) from drawing by detecting:
        1. Circle and diamond vector shapes near dimensions
        2. Characteristic index letters (A-Z) enclosed in circles or adjacent to diamonds
        3. Special flag notes specifying SC/CTQ requirements (e.g. Note 12, Note 13)
        4. Matches each SC feature with its nearby dimension callout
        """
        import math
        sc_features = []
        try:
            import pymupdf
            doc = pymupdf.open(self.pdf_path)
            for page_idx, page in enumerate(doc):
                drawings = page.get_drawings()
                words = page.get_text("words")
                
                circle_centers = []
                diamond_centers = []
                v_lines = []
                h_lines = []
                diag_lines = []

                for d in drawings:
                    r = d["rect"]
                    w, h = r.width, r.height
                    cx = (r.x0 + r.x1) / 2.0
                    cy = (r.y0 + r.y1) / 2.0
                    if 10 <= w <= 28 and 10 <= h <= 28 and abs(w - h) <= 3:
                        circle_centers.append((cx, cy))
                    elif 7 <= w <= 32 and 7 <= h <= 32 and abs(w - h) <= 8:
                        diamond_centers.append((cx, cy))

                    if w <= 1.0 and 10 <= h <= 30:
                        v_lines.append((cx, cy, r))
                    elif h <= 1.0 and 8 <= w <= 25:
                        h_lines.append((cx, cy, r))
                    elif 3 <= w <= 15 and 5 <= h <= 15:
                        diag_lines.append((cx, cy, r))

                # 1. Detect CTQ cross-diamonds (perpendicular cross + diagonal envelope, e.g. Note 7 CTQ symbol)
                ctq_diamonds = []
                for vx, vy, vr in v_lines:
                    for hx, hy, hr in h_lines:
                        if abs(vx - hx) < 2.0 and abs(vy - hy) < 2.0:
                            cx, cy = (vx + hx) / 2.0, (vy + hy) / 2.0
                            diags_near = [d for d in diag_lines if math.hypot(d[0]-cx, d[1]-cy) < 12.0]
                            if len(diags_near) >= 3:
                                ctq_diamonds.append((cx, cy))

                for dcx, dcy in ctq_diamonds:
                    if dcx < 150: # Skip Note 7 legend in margin notes
                        continue
                    nearby_nums = []
                    for w in words:
                        txt = w[4].strip()
                        if re.match(r'^(\.|\d+\.)\d+$', txt):
                            wx = (w[0] + w[2]) / 2.0
                            wy = (w[1] + w[3]) / 2.0
                            dist = math.hypot(wx - dcx, wy - dcy)
                            if dist < 65:
                                nearby_nums.append((dist, float(txt), txt, wx, wy))
                    nearby_nums.sort(key=lambda x: x[0])
                    if nearby_nums:
                        best = nearby_nums[0]
                        dim_x, dim_y = best[3], best[4]
                        nearby_tokens = [w[4].strip() for w in words if math.hypot((w[0]+w[2])/2 - dim_x, (w[1]+w[3])/2 - dim_y) < 45]
                        nearby_tol_str = " ".join(nearby_tokens)
                        sc_features.append({
                            "type": "CTQ",
                            "letter": f"CTQ-{best[2]}",
                            "nominal": best[1],
                            "raw_text": best[2],
                            "nearby_tol_str": nearby_tol_str,
                            "page": page_idx + 1,
                            "req_cpk": 1.33,
                            "pos": (round(dcx, 1), round(dcy, 1))
                        })

                # 2. Detect single uppercase letters inside circle in drawing body
                sc_letters = {}
                for w in words:
                    txt = w[4].strip()
                    if len(txt) == 1 and txt.isupper() and txt not in ["R", "X", "O"]:
                        wx = (w[0] + w[2]) / 2.0
                        wy = (w[1] + w[3]) / 2.0
                        circ_dist = min([math.hypot(wx - cx, wy - cy) for cx, cy in circle_centers]) if circle_centers else 999
                        if circ_dist < 5.0 and wx > 650:
                            sc_letters[txt] = (wx, wy)
                                    
                # For each letter, find closest dimension callout and tolerance tokens
                for letter, (lx, ly) in sorted(sc_letters.items()):
                    nums = []
                    for w in words:
                        txt = w[4].strip()
                        if re.match(r'^\d+\.\d+$', txt):
                            wx = (w[0] + w[2]) / 2.0
                            wy = (w[1] + w[3]) / 2.0
                            dist = math.hypot(wx - lx, wy - ly)
                            if dist < 45:
                                nums.append((dist, float(txt), txt, wx, wy))
                    nums.sort(key=lambda x: x[0])
                    if not nums:
                        for w in words:
                            txt = w[4].strip().replace('(', '').replace(')', '')
                            if re.match(r'^\d+$', txt) and int(txt) not in [2, 4]:
                                wx = (w[0] + w[2]) / 2.0
                                wy = (w[1] + w[3]) / 2.0
                                dist = math.hypot(wx - lx, wy - ly)
                                if dist < 45:
                                    nums.append((dist, float(txt), txt, wx, wy))
                        nums.sort(key=lambda x: x[0])
                    
                    best_dim = None
                    raw_text = ""
                    nearby_tol_str = ""
                    if nums:
                        best_dim = nums[0][1]
                        raw_text = nums[0][2]
                        dim_x, dim_y = nums[0][3], nums[0][4]
                        nearby_tokens = [w[4].strip() for w in words if math.hypot((w[0]+w[2])/2 - dim_x, (w[1]+w[3])/2 - dim_y) < 45]
                        nearby_tol_str = " ".join(nearby_tokens)

                    sc_features.append({
                        "type": "SC",
                        "letter": letter,
                        "nominal": best_dim,
                        "raw_text": raw_text,
                        "nearby_tol_str": nearby_tol_str,
                        "page": page_idx + 1,
                        "pos": (round(lx, 1), round(ly, 1))
                    })
                    
                full_text = page.get_text()
                m_pin = re.search(r'MAINTAIN\s*([\d\.]+)\s*(?:MM)?\s*TO\s*([\d\.]+)\s*(?:MM)?.*?PIN\s+GAUGE', full_text, re.IGNORECASE | re.DOTALL)
                if m_pin:
                    sc_features.append({
                        "type": "SC",
                        "letter": "Note 13",
                        "nominal": float(m_pin.group(1)),
                        "lower_bound": float(m_pin.group(2)),
                        "raw_text": m_pin.group(0)[:80],
                        "page": page_idx + 1,
                        "special": "Pin Gauge"
                    })
        except Exception:
            pass
        return sc_features

    def extract_dimensions_and_tolerances(self):
        text = " ".join(self.pages_text)
        
        # 1. Parse generic tolerance from notes (e.g. TOLERANCE: +/- 1/16)
        gen_tol = 0.0625  # default (1/16)
        tol_match = re.search(r'TOLERANCE:\s*(?:\+/-|±|#)?\s*(\d+/\d+|\d*\.\d+|\d+)', text, re.IGNORECASE)
        if tol_match:
            try:
                gen_tol = parse_fraction(tol_match.group(1))
            except Exception:
                pass
                
        # 1.1 Parse ASME Y14.5M style tolerances (.XX +/- .01, .XXX +/- .005)
        tol_2_place = None
        tol_3_place = None
        m_xx = re.search(r'\.XX\b[^\d\n]*?(\.\d+|\d+\.\d+)', text, re.IGNORECASE)
        if m_xx:
            try:
                tol_2_place = float(m_xx.group(1))
            except Exception:
                pass
        m_xxx = re.search(r'\.XXX\b[^\d\n]*?(\.\d+|\d+\.\d+)', text, re.IGNORECASE)
        if m_xxx:
            try:
                tol_3_place = float(m_xxx.group(1))
            except Exception:
                pass
        if gen_tol == 0.0625 and tol_3_place is not None:
            gen_tol = tol_3_place
                
        ang_tol = 1.0
        ang_match = re.search(r'ANGLES:\s*(?:\+/-|±|#)?\s*(\d+(?:\.\d+)?)', text, re.IGNORECASE)
        if ang_match:
            try:
                ang_tol = float(ang_match.group(1))
            except Exception:
                pass
                
        # Clean brackets: [ ] -> []
        clean_text = re.sub(r'\[\s*\]', '[]', text)
        norm_text = re.sub(r'\s+', ' ', clean_text)
        norm_text = re.sub(r'(\d+)\s*-\s*(\d+)', r'\1-\2', norm_text)
        norm_text = re.sub(r'(\d+)\s*/\s*(\d+)', r'\1/\2', norm_text)
        norm_text = re.sub(r'\[\s*([\d\s,\.]+)\s*\]', lambda m: "[" + m.group(1).replace(" ", "") + "]", norm_text)
        
        candidates = []
        tokens = norm_text.split()
        for i, tok in enumerate(tokens):
            tok_clean = tok.strip().replace(',', '.')
            if re.search(r'\d', tok_clean):
                est_s = re.sub(r'^[RrQq\-\+]+', '', tok_clean)
                est_s = re.sub(r'^\d+-(?!\d+/\d+)', '', est_s)
                if est_s.startswith('[') and est_s.endswith(']'):
                    est_s = est_s[1:-1].strip()
                est_s = re.sub(r'\[\s*\]$', '', est_s)
                est_s = re.sub(r'\[\]$', '', est_s)
                
                # Clean merge artifacts (e.g. 3-7/81 -> 3-7/8)
                m_frac = re.match(r'^(\d+)-(\d+)/(\d+)$', est_s)
                if m_frac:
                    whole, num, den = m_frac.groups()
                    for std_den in ['16', '32', '64', '2', '4', '8']:
                        if den.startswith(std_den) and len(den) > len(std_den):
                            est_s = f"{whole}-{num}/{std_den}"
                            break
                else:
                    m_frac2 = re.match(r'^(\d+)/(\d+)$', est_s)
                    if m_frac2:
                        num, den = m_frac2.groups()
                        for std_den in ['16', '32', '64', '2', '4', '8']:
                            if den.startswith(std_den) and len(den) > len(std_den):
                                est_s = f"{num}/{std_den}"
                                break

                # Support leading dot (e.g. .43, .032, .75)
                float_match = re.search(r'(\d+-\d+/\d+|\d+/\d+|\d*\.\d+|\d+)', est_s)
                if float_match and float_match.group(1):
                    try:
                        val_str = float_match.group(1)
                        if '/' in val_str:
                            val = parse_fraction(val_str)
                        else:
                            val = float(val_str)
                        candidates.append({
                            "idx": i,
                            "raw_str": tok,
                            "clean_str": tok_clean,
                            "val": val
                        })
                    except Exception:
                        pass
                        
        candidates.sort(key=lambda x: x["idx"])
        dimensions = []
        used = set()
        
        for i, c1 in enumerate(candidates):
            if c1["idx"] in used:
                continue
            # Restrict distance to at most i+2 to avoid false pairings
            for j in range(i + 1, min(i + 3, len(candidates))):
                c2 = candidates[j]
                if c2["idx"] in used:
                    continue
                v1 = c1["val"]
                v2 = c2["val"]
                if v1 == 0 or v2 == 0:
                    continue
                ratio1 = v1 / v2
                ratio2 = v2 / v1
                is_match = False
                inch_val = None
                mm_val = None
                inch_token = None
                mm_token = None
                
                has_metric_indicator = ('[' in c1["raw_str"] or ']' in c1["raw_str"] or
                                        '[' in c2["raw_str"] or ']' in c2["raw_str"] or
                                        'mm' in c1["raw_str"].lower() or 'mm' in c2["raw_str"].lower())
                
                if (abs(ratio1 - 25.4) < 0.8) and has_metric_indicator:
                    is_match = True
                    inch_val = v2
                    mm_val = v1
                    inch_token = c2
                    mm_token = c1
                elif (abs(ratio2 - 25.4) < 0.8) and has_metric_indicator:
                    is_match = True
                    inch_val = v1
                    mm_val = v2
                    inch_token = c1
                    mm_token = c2
                    
                if is_match:
                    inch_nom, inch_tol = parse_nom_tol(inch_token["clean_str"], inch_val)
                    mm_nom, mm_tol = parse_nom_tol(mm_token["clean_str"], mm_val)
                    if inch_nom is None:
                        inch_nom = inch_val
                    if mm_nom is None:
                        mm_nom = mm_val
                    dimensions.append({
                        "inch_str": str(inch_nom),
                        "inch_val": inch_nom,
                        "mm_str": str(mm_nom),
                        "mm_val": mm_nom,
                        "inch_tol": inch_tol,
                        "mm_tol": mm_tol,
                        "type": "paired_smart"
                    })
                    used.add(c1["idx"])
                    used.add(c2["idx"])
                    break
                    
        # Always add all other unmatched candidates as single dimensions
        for c in candidates:
            if c["idx"] not in used:
                if c["clean_str"] not in ["1/16", "1/1", "0.037", ".02", "0.02", "0,02", ".05", "0.05", "0,05", ".005", "0.005", "0,005"]:
                    dimensions.append({
                        "inch_str": c["clean_str"],
                        "inch_val": c["val"],
                        "mm_str": c["clean_str"],
                        "mm_val": c["val"],
                        "inch_tol": None,
                        "mm_tol": None,
                        "type": "single"
                    })
                    
        # Dedicate a second pass to scan for explicit tolerances
        dimensions = self.extract_individual_tolerances(text, dimensions)

        return {
            "general_tolerance": gen_tol,
            "tol_2_place": tol_2_place,
            "tol_3_place": tol_3_place,
            "angle_tolerance": ang_tol,
            "dimensions": dimensions,
            "text": text
        }

    def extract_individual_tolerances(self, text, dimensions):
        tol_pattern = re.compile(
            r'^\s*(?:\]|\)|\]\s*|\)\s*)?(?:'
            r'(?:±|\+/-|#)\s*(\d*\.\d+|\d+)[\$°]?'
            r'|(?:\+\s*(\d*\.\d+|\d+)[\$°]?\s*(?:/|\s+)\s*-\s*(\d*\.\d+|\d+)[\$°]?)'
            r'|(?:\+\s*(\d*\.\d+|\d+)[\$°]?\s*(?:/|\s+)\s*0[\$°]?)'
            r'|(?:0[\$°]?\s*(?:/|\s+)\s*-\s*(\d*\.\d+|\d+)[\$°]?)'
            r'|(?:\s*(\.\d{2,4}))'
            r')', re.IGNORECASE
        )
        ang_tol_pattern = re.compile(
            r'^\s*[\x01\x02\$°]+\s*(?:(?:±|\+/-|#)\s*(\d*\.\d+|\d+)[\x01\x02\$°]?|(\d*\.\d+|\d+)[\x01\x02\$°])(?![xX\w])',
            re.IGNORECASE
        )
        for entry in dimensions:
            mm_tol_up = mm_tol_lo = None
            inch_tol_up = inch_tol_lo = None

            # 0. Check if entry string itself contains angle dimension and tolerance (e.g. 50$5$ or 50°±5°)
            raw_s = entry.get("mm_str") or entry.get("inch_str") or ""
            m_self_ang = re.match(r'^(\d+(?:\.\d+)?)[^\d\w]*[\$°][^\d\w]*(\d+(?:\.\d+)?)[^\d\w]*[\$°]?[^\d\w]*$', raw_s)
            if m_self_ang:
                nom_ang = float(m_self_ang.group(1))
                tol_ang = float(m_self_ang.group(2))
                entry["inch_val"] = nom_ang
                entry["mm_val"] = nom_ang
                entry["inch_str"] = f"{nom_ang}°"
                entry["mm_str"] = f"{nom_ang}°"
                mm_tol_up = mm_tol_lo = tol_ang
                inch_tol_up = inch_tol_lo = tol_ang
                entry["is_angle"] = True

            # 1. Search next to mm_str if available
            mm_str = entry.get("mm_str")
            if mm_str and mm_tol_up is None:
                clean_target = re.sub(r'^\d+[-\sXx]+', '', mm_str)
                pattern = re.compile(rf'(?<![\d\.])' + re.escape(clean_target) + rf'(?![\d\.])')
                for match in pattern.finditer(text):
                    sub = text[match.end():match.end() + 25]
                    m_tol = tol_pattern.match(sub)
                    if m_tol:
                        if m_tol.group(1): # Symmetric
                            val = float(m_tol.group(1))
                            mm_tol_up = mm_tol_lo = val
                        elif m_tol.group(2) and m_tol.group(3): # Asymmetric +X / -Y
                            mm_tol_up = float(m_tol.group(2))
                            mm_tol_lo = float(m_tol.group(3))
                        elif m_tol.group(4): # Asymmetric +X / 0
                            mm_tol_up = float(m_tol.group(4))
                            mm_tol_lo = 0.0
                        elif m_tol.group(5): # Asymmetric 0 / -Y
                            mm_tol_up = 0.0
                            mm_tol_lo = float(m_tol.group(5))
                        elif len(m_tol.groups()) >= 6 and m_tol.group(6):
                            val = float(m_tol.group(6))
                            if val <= 0.03 and not (entry.get("is_angle") or "$" in str(mm_str) or "°" in str(mm_str)):
                                mm_tol_up = mm_tol_lo = val
                        break # Use the first found annotated tolerance
                    else:
                        m_ang_tol = ang_tol_pattern.match(sub)
                        if m_ang_tol:
                            val = float(m_ang_tol.group(1) or m_ang_tol.group(2))
                            mm_tol_up = mm_tol_lo = val
                            inch_tol_up = inch_tol_lo = val
                            entry["is_angle"] = True
                            break

            # 2. Search next to inch_str if available
            inch_str = entry.get("inch_str")
            if inch_str and inch_tol_up is None:
                clean_target = re.sub(r'^\d+[-\sXx]+', '', inch_str)
                pattern = re.compile(rf'(?<![\d\.])' + re.escape(clean_target) + rf'(?![\d\.])')
                for match in pattern.finditer(text):
                    sub = text[match.end():match.end() + 25]
                    m_tol = tol_pattern.match(sub)
                    if m_tol:
                        if m_tol.group(1): # Symmetric
                            val = float(m_tol.group(1))
                            inch_tol_up = inch_tol_lo = val
                        elif m_tol.group(2) and m_tol.group(3): # Asymmetric
                            inch_tol_up = float(m_tol.group(2))
                            inch_tol_lo = float(m_tol.group(3))
                        elif m_tol.group(4): # Asymmetric +X / 0
                            inch_tol_up = float(m_tol.group(4))
                            inch_tol_lo = 0.0
                        elif m_tol.group(5): # Asymmetric 0 / -Y
                            inch_tol_up = 0.0
                            inch_tol_lo = float(m_tol.group(5))
                        elif len(m_tol.groups()) >= 6 and m_tol.group(6):
                            val = float(m_tol.group(6))
                            if val <= 0.03 and not (entry.get("is_angle") or "$" in str(inch_str) or "°" in str(inch_str)):
                                inch_tol_up = inch_tol_lo = val
                        break
                    else:
                        m_ang_tol = ang_tol_pattern.match(sub)
                        if m_ang_tol:
                            val = float(m_ang_tol.group(1) or m_ang_tol.group(2))
                            inch_tol_up = inch_tol_lo = val
                            mm_tol_up = mm_tol_lo = val
                            entry["is_angle"] = True
                            break

            # 3. Cross-convert tolerances if only one unit was explicitly annotated
            if mm_tol_up is not None and inch_tol_up is None:
                inch_tol_up = mm_tol_up / 25.4
                inch_tol_lo = mm_tol_lo / 25.4
            elif inch_tol_up is not None and mm_tol_up is None:
                mm_tol_up = inch_tol_up * 25.4
                mm_tol_lo = inch_tol_lo * 25.4

            if entry.get("inch_tol") is not None and inch_tol_up is None:
                inch_tol_up = inch_tol_lo = entry["inch_tol"]
            if entry.get("mm_tol") is not None and mm_tol_up is None:
                mm_tol_up = mm_tol_lo = entry["mm_tol"]

            entry["mm_tol_up"] = mm_tol_up
            entry["mm_tol_lo"] = mm_tol_lo
            entry["inch_tol_up"] = inch_tol_up
            entry["inch_tol_lo"] = inch_tol_lo

        return dimensions

def clean_val(v):
    """Normalize a string value for comparison: strip, uppercase, unify dashes and O/0."""
    if v is None:
        return ""
    s = str(v).strip().upper()
    # Replace Unicode dash/hyphen variants with standard ASCII hyphen
    s = re.sub(r'[\u2010-\u2015\u2212\uFE58\uFE63\uFF0D]', '-', s)
    # Remove any remaining invisible/non-printable characters
    s = re.sub(r'[^\x20-\x7E]', '', s)
    # Collapse multiple spaces
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def parse_fcc_metadata(text):
    meta = {}
    ignore_words = {"LABEL", "FOR", "THE", "SELECT", "PRODUCTS", "TO", "A", "OF", "IS", "AND", "OR", "ITS", "WHERE"}
    
    # 1. Model
    model_matches = re.finditer(r'\bM[O0]DEL\s*[:\-\s]?\s*([A-Z0-9a-z\-]+)\b', text, re.IGNORECASE)
    for m in model_matches:
        val = m.group(1).upper()
        if val not in ignore_words:
            meta["model"] = val
            break
            
    # 2. FCC ID
    fcc_matches = re.finditer(r'\bFCC\s*I[Dd]\s*[:\-\s]?\s*([A-Z0-9a-z\-]+)\b', text, re.IGNORECASE)
    for m in fcc_matches:
        val = m.group(1).upper()
        if val not in ignore_words:
            meta["fcc_id"] = val
            break
            
    # 3. IC
    ic_matches = re.finditer(r'\bIC\s*[:\-\s]?\s*([A-Z0-9a-z\-]+)\b', text, re.IGNORECASE)
    for m in ic_matches:
        val = m.group(1).upper()
        if val not in ignore_words and val not in {"REGISTRATION", "CERTIFICATION", "NUMBER", "INDUSTRY", "CANADA"}:
            meta["ic"] = val
            break
            
    # 4. HVIN
    hvin_matches = re.finditer(r'\bHVIN\s*[:\-\s]?\s*([A-Z0-9a-z\-]+)\b', text, re.IGNORECASE)
    for m in hvin_matches:
        val = m.group(1).upper()
        if val not in ignore_words:
            meta["hvin"] = val
            break
            
    return meta


def audit_fff_table_and_content(wb, dwg_spec_data=None, dwg_text="", dim_sample_count=None):
    """
    Audits active visible 3F/FFF worksheets for:
    1. Copy-paste residue in Requirement / Test Performed (e.g. foreign project names like 'FOSTER' vs drawing 'TITAN')
    2. Sample count discrepancies: 'Qty Tested' (e.g. 5 pcs) vs 100% Dimensional measured sample count (e.g. 3 pcs)
    3. Generic verification descriptions without specific drawing criteria
    4. Disposition compliance (must be 'Accepted' / 'Approved' / 'Pass')
    """
    results = []
    KNOWN_PROJECTS = [
        "TITAN", "FOSTER", "HALO", "AURA", "PREMIS", "OBSIDIAN", "POWERBOLT", 
        "SMARTCODE", "KEYWAY", "TOUCHSCREEN", "DEADBOLT", "INTERCONNECT", 
        "CABRIO", "AVENUE", "VERONA", "SEVILLE", "MONTARA", "HAWTHORNE", 
        "SAN CLEMENTE", "AMHERST", "TUSTIN", "LADERA", "MILAN", "HALIFAX", 
        "VEDANI", "SYDNEY", "PEYTON", "BALBOA", "COMMONWEALTH", "CAMINO", 
        "DELANEY", "TELL", "DE2000"
    ]
    
    dwg_upper = (str(dwg_text or '') + " " + json.dumps(dwg_spec_data or {})).upper()
    valid_dwg_projects = {p for p in KNOWN_PROJECTS if p in dwg_upper}
    
    # Auto-infer dimensional sample count if not provided
    if dim_sample_count is None and wb is not None:
        dim_sheets = [s for s in wb.sheetnames if "Dimensional" in s and wb[s].sheet_state == 'visible']
        max_s = 0
        for ds in dim_sheets:
            d_ws = wb[ds]
            cols = find_dimensional_columns(d_ws)
            for r in range(25, min(d_ws.max_row + 1, 100)):
                s_val = d_ws.cell(row=r, column=cols["sample"]).value
                if s_val is not None:
                    try:
                        s_int = int(float(str(s_val).strip()))
                        if s_int > max_s:
                            max_s = s_int
                    except Exception:
                        pass
        if max_s > 0:
            dim_sample_count = max_s

    fff_sheets = [s for s in wb.sheetnames if ("FFF" in s.upper() or "3F" in s.upper() or "FIT & FUNCTION" in s.upper()) and wb[s].sheet_state == 'visible']
    for sname in fff_sheets:
        ws = wb[sname]
        col_map = {"verif": 1, "req": 6, "qty": 16, "disp": 19}
        # Find headers dynamically
        for r in range(15, min(30, ws.max_row + 1)):
            for c in range(1, min(25, ws.max_column + 1)):
                val = str(ws.cell(row=r, column=c).value or '').strip().lower()
                if 'verification' in val:
                    col_map['verif'] = c
                elif 'requirement' in val or 'test performed' in val:
                    col_map['req'] = c
                elif 'qty' in val or 'quantity' in val:
                    col_map['qty'] = c
                elif 'disposition' in val:
                    col_map['disp'] = c
                    
        for r in range(20, min(60, ws.max_row + 1)):
            verif_val = str(ws.cell(row=r, column=col_map['verif']).value or '').strip()
            if verif_val.lower() in ['form', 'fit', 'function']:
                req_val = str(ws.cell(row=r, column=col_map['req']).value or '').strip()
                qty_val = ws.cell(row=r, column=col_map['qty']).value
                disp_val = str(ws.cell(row=r, column=col_map['disp']).value or '').strip()
                
                # Check 1: Project copy-paste check
                req_upper = req_val.upper()
                for p in KNOWN_PROJECTS:
                    if p in req_upper:
                        if valid_dwg_projects and p not in valid_dwg_projects:
                            results.append({
                                'status': 'FAIL',
                                'sheet': sname.strip(),
                                'row': r,
                                'field': f'{verif_val} (Requirement / Test Performed)',
                                'message': f"工作表 [{sname.strip()}] 第 {r} 列 {verif_val} 欄位填寫 '{req_val}'，偵測到舊專案名稱 '{p}'，但本圖面為 '{list(valid_dwg_projects)[0]}'！德昌/供應商疑似自舊專案 PQP 複製貼上未修改！"
                            })
                            
                # Check 2: Sample count alignment
                if qty_val is not None:
                    try:
                        q_int = int(float(str(qty_val).strip()))
                        if dim_sample_count and q_int != dim_sample_count:
                            results.append({
                                'status': 'WARNING',
                                'sheet': sname.strip(),
                                'row': r,
                                'field': f'{verif_val} (Qty Tested)',
                                'message': f"工作表 [{sname.strip()}] 第 {r} 列 {verif_val} 之測試數量 (Qty Tested) 填寫為 {q_int} 顆，但 100% 尺寸量測表中僅量測了 {dim_sample_count} 顆樣品，兩表數量自相矛盾！"
                            })
                    except Exception:
                        pass
                        
                # Check 3: Generic description
                if req_val.lower() in ['comply with the drawing', 'comply with drawing', 'per drawing']:
                    results.append({
                        'status': 'INFO',
                        'sheet': sname.strip(),
                        'row': r,
                        'field': f'{verif_val} (Requirement)',
                        'message': f"工作表 [{sname.strip()}] 第 {r} 列 {verif_val} 檢驗條件過於籠統 ('{req_val}')。建議依圖面規範（如顏色 White、字體 Bahnschrift、標籤材質）明確列出具體驗收標準。"
                    })
                elif req_val.lower() in ['fit well and no interference', 'no interference', 'fit well']:
                    results.append({
                        'status': 'INFO',
                        'sheet': sname.strip(),
                        'row': r,
                        'field': f'{verif_val} (Requirement)',
                        'message': f"工作表 [{sname.strip()}] 第 {r} 列 {verif_val} 檢驗條件過於籠統 ('{req_val}')。建議補充組裝至鎖體時之貼合度與無干涉判定標準。"
                    })
                    
                # Check 4: Disposition
                if disp_val and disp_val.lower() not in ['accepted', 'acceptable', 'approved', 'pass']:
                    results.append({
                        'status': 'FAIL',
                        'sheet': sname.strip(),
                        'row': r,
                        'field': f'{verif_val} (Disposition)',
                        'message': f"工作表 [{sname.strip()}] 第 {r} 列 {verif_val} 判定狀態為 '{disp_val}'，非合規判定狀態 (Accepted)！"
                    })
    return results


def run_fff_photo_audit(pqp_path, dwg_text, wb=None, dwg_spec_data=None, dim_sample_count=None):
    import io
    import zipfile
    import xml.etree.ElementTree as ET
    from PIL import Image
    
    dwg_meta = parse_fcc_metadata(dwg_text)
    
    table_findings = []
    if wb is not None:
        table_findings = audit_fff_table_and_content(wb, dwg_spec_data=dwg_spec_data, dwg_text=dwg_text, dim_sample_count=dim_sample_count)
    
    # 1. Check for embedded Excel workbooks in FFF sheet (e.g. xl/embeddings/Microsoft_Excel_Worksheet*.xlsx)
    embedded_fff_info = []
    try:
        with zipfile.ZipFile(io.BytesIO(read_file_bytes_safe(pqp_path)), 'r') as z:
            wb_xml = z.read('xl/workbook.xml')
            wb_tree = ET.fromstring(wb_xml)
            sheets = {}
            for sheet in wb_tree.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
                sheets[sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']] = sheet.attrib['name']
            
            rels_xml = z.read('xl/_rels/workbook.xml.rels')
            rels_tree = ET.fromstring(rels_xml)
            sheet_targets = {}
            for rel in rels_tree.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                if rel.attrib['Id'] in sheets:
                    sheet_targets[sheets[rel.attrib['Id']]] = rel.attrib['Target']
            
            fff_embedded_files = []
            for sheet_name, target in sheet_targets.items():
                s_lower = sheet_name.lower().strip()
                if 'fff' in s_lower or s_lower == '3f' or 'fit & function' in s_lower:
                    sheet_filename = target.split('/')[-1]
                    sheet_rels_path = f'xl/worksheets/_rels/{sheet_filename}.rels'
                    if sheet_rels_path in z.namelist():
                        s_rels_tree = ET.fromstring(z.read(sheet_rels_path))
                        for rel in s_rels_tree.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            t = rel.attrib['Target']
                            if 'embedding' in t.lower() and (t.endswith('.xlsx') or t.endswith('.xls')):
                                emb_path = 'xl/' + t.lstrip('../') if not t.startswith('xl/') else t
                                if emb_path in z.namelist():
                                    fff_embedded_files.append(emb_path)
            
            for emb_path in fff_embedded_files:
                emb_bytes = z.read(emb_path)
                import tempfile
                with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                    tmp_emb_path = tmp.name
                try:
                    with open(tmp_emb_path, 'wb') as tmp_f:
                        tmp_f.write(emb_bytes)
                    emb_wb = openpyxl.load_workbook(tmp_emb_path, data_only=True)
                    for sname in emb_wb.sheetnames:
                        ws_emb = emb_wb[sname]
                        img_count = len(getattr(ws_emb, '_images', []))
                        if img_count == 0:
                            with zipfile.ZipFile(tmp_emb_path, 'r') as ez:
                                img_count = len([n for n in ez.namelist() if n.startswith('xl/media/') and not n.endswith('.emf')])
                        
                        form_status = None
                        fit_status = None
                        func_status = None
                        max_r = min(ws_emb.max_row, 200)
                        max_c = min(ws_emb.max_column, 30)
                        for r in range(1, max_r + 1):
                            row_txt = " ".join([str(ws_emb.cell(row=r, column=c).value or '') for c in range(1, max_c + 1)])
                            if 'FORM' in row_txt.upper() and ('ACCEPTED' in row_txt.upper() or 'PASS' in row_txt.upper()):
                                form_status = 'Accepted'
                            if 'FIT' in row_txt.upper() and ('ACCEPTED' in row_txt.upper() or 'PASS' in row_txt.upper()):
                                fit_status = 'Accepted'
                            if 'FUNCTION' in row_txt.upper() and ('ACCEPTED' in row_txt.upper() or 'PASS' in row_txt.upper()):
                                func_status = 'Accepted'

                        embedded_fff_info.append({
                            "sheet_name": sname,
                            "file_path": emb_path,
                            "image_count": img_count,
                            "form_status": form_status,
                            "fit_status": fit_status,
                            "func_status": func_status
                        })
                finally:
                    if os.path.exists(tmp_emb_path):
                        os.remove(tmp_emb_path)
    except Exception as e:
        pass

    if embedded_fff_info:
        valid_reports = [e for e in embedded_fff_info if e["image_count"] > 0 or (e["form_status"] and e["fit_status"] and e["func_status"])]
        if valid_reports:
            results = list(table_findings)
            for r in valid_reports:
                results.append({
                    "status": "PASS",
                    "message": f"嵌入的 FFF 工作表 [{r['sheet_name'].strip()}] 驗證通過：包含 {r['image_count']} 張實體照片與 Form/Fit/Function 簽核紀錄。"
                })
            return results

    # 2. Extract photos directly from FFF worksheet in zip or ws._images
    extracted_images = []
    try:
        with zipfile.ZipFile(io.BytesIO(read_file_bytes_safe(pqp_path)), 'r') as z:
            wb_xml = z.read('xl/workbook.xml')
            wb_tree = ET.fromstring(wb_xml)
            ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            rel_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            p_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
            
            fff_sheet_rids = []
            for sheet in wb_tree.findall(f'{{{ns}}}sheets/{{{ns}}}sheet'):
                s_name = sheet.attrib.get('name', '').lower().strip()
                if 'fff' in s_name or s_name == '3f' or 'fit & function' in s_name:
                    fff_sheet_rids.append(sheet.attrib.get(f'{{{rel_ns}}}id'))
            
            wb_rels_xml = z.read('xl/_rels/workbook.xml.rels')
            wb_rels_tree = ET.fromstring(wb_rels_xml)
            
            fff_sheet_targets = []
            for rel in wb_rels_tree.findall(f'{{{p_ns}}}Relationship'):
                if rel.attrib.get('Id') in fff_sheet_rids:
                    fff_sheet_targets.append(rel.attrib.get('Target'))
            
            for target in fff_sheet_targets:
                sheet_filename = target.split('/')[-1]
                sheet_rels_path = f'xl/worksheets/_rels/{sheet_filename}.rels'
                if sheet_rels_path in z.namelist():
                    s_rels_tree = ET.fromstring(z.read(sheet_rels_path))
                    for rel in s_rels_tree.findall(f'{{{p_ns}}}Relationship'):
                        if 'drawing' in rel.attrib.get('Type', '').lower():
                            drawing_target = rel.attrib.get('Target')
                            drawing_filename = drawing_target.split('/')[-1]
                            drawing_rels_path = f'xl/drawings/_rels/{drawing_filename}.rels'
                            if drawing_rels_path in z.namelist():
                                d_rels_tree = ET.fromstring(z.read(drawing_rels_path))
                                for d_rel in d_rels_tree.findall(f'{{{p_ns}}}Relationship'):
                                    if 'image' in d_rel.attrib.get('Type', '').lower():
                                        img_target = d_rel.attrib.get('Target')
                                        img_zip_path = 'xl/' + img_target.lstrip('../') if not img_target.startswith('xl/') else img_target
                                        if img_zip_path in z.namelist() and not img_zip_path.endswith('.emf'):
                                            img_bytes = z.read(img_zip_path)
                                            if len(img_bytes) > 20000:
                                                extracted_images.append(img_bytes)
    except Exception:
        pass

    # Fallback to ws._images
    if not extracted_images:
        try:
            wb = openpyxl.load_workbook(pqp_path)
            fff_sheet_names = [s for s in wb.sheetnames if "FFF" in s.upper() or "3F" in s.upper()]
            if fff_sheet_names:
                ws = wb[fff_sheet_names[0]]
                if hasattr(ws, '_images') and ws._images:
                    for img in ws._images:
                        if hasattr(img, 'ref') and isinstance(img.ref, bytes):
                            extracted_images.append(img.ref)
        except Exception:
            pass

    if not extracted_images:
        if table_findings:
            return table_findings
        return [{"status": "WARNING", "message": "未在 FFF 工作表中找到任何實體照片。"}]

    # Skip OCR if no drawing text is available (PQP-only audit mode)
    if not dwg_text or not dwg_text.strip():
        res = list(table_findings)
        res.append({"status": "INFO", "message": f"已找到 {len(extracted_images)} 張 FFF 照片，但無工程圖面資料可供比對，跳過照片 OCR 驗證。"})
        return res

    try:
        from rapidocr_onnxruntime import RapidOCR
        engine = RapidOCR()
    except ImportError:
        res = list(table_findings)
        res.append({"status": "WARNING", "message": "未安裝 rapidocr-onnxruntime，跳過 3F 照片比對。"})
        return res

    # Build drawing keywords set for intelligent matching
    dwg_lower = dwg_text.lower()
    has_fcc_in_dwg = bool(dwg_meta.get("fcc_id") or dwg_meta.get("ic") or dwg_meta.get("hvin"))
    
    # Common brand and label keywords
    known_brands = ["kwikset", "weiser", "baldwin", "yale", "august", "assa abloy"]
    matched_brands_in_dwg = [b for b in known_brands if b in dwg_lower]
    
    mismatches = []
    photo_count = 0
    
    # Limit OCR to max 8 most representative photos if there are dozens
    images_to_ocr = extracted_images[:8] if len(extracted_images) > 8 else extracted_images
    
    for idx, img_bytes in enumerate(images_to_ocr):
        try:
            pil_img = Image.open(io.BytesIO(img_bytes))
            photo_count += 1
            if max(pil_img.size) > 1000:
                pil_img.thumbnail((1000, 1000))
            
            import tempfile
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                tmp_img_path = tmp.name
            try:
                pil_img.save(tmp_img_path)
                ocr_results, elapse = engine(tmp_img_path)
            finally:
                if os.path.exists(tmp_img_path):
                    os.remove(tmp_img_path)
                    
            if not ocr_results:
                mismatches.append({
                    "status": "WARNING",
                    "message": f"無法識別 3F/FFF 照片 {photo_count} 中的文字。"
                })
                continue
                
            ocr_text = " ".join([res[1] for res in ocr_results])
            ocr_lower = ocr_text.lower()
            photo_meta = parse_fcc_metadata(ocr_text)
            
            if has_fcc_in_dwg:
                # Compare FCC/IC parameters
                keys_to_compare = {
                    "model": "MODEL",
                    "fcc_id": "FCC ID",
                    "ic": "IC",
                    "hvin": "HVIN"
                }
                for key, label in keys_to_compare.items():
                    dwg_val = dwg_meta.get(key)
                    photo_val = photo_meta.get(key)
                    dwg_clean = clean_val(dwg_val)
                    photo_clean = clean_val(photo_val)
                    if dwg_clean and photo_clean:
                        if dwg_clean != photo_clean:
                            mismatches.append({
                                "status": "FAIL",
                                "field": label,
                                "expected": dwg_clean,
                                "found": photo_clean,
                                "photo_num": photo_count
                            })
                    elif dwg_clean and not photo_clean:
                        mismatches.append({
                            "status": "WARNING",
                            "field": label,
                            "expected": dwg_clean,
                            "found": "未找到",
                            "photo_num": photo_count
                        })
            else:
                # Intelligent Setup Code / Support / Model Label Matching
                matched_features = []
                for brand in known_brands:
                    if brand in ocr_lower:
                        matched_features.append(f"Brand: {brand.title()}")
                        
                for phrase in ["wi-fi smart lock", "smart lock", "setup code", "scan here", "instruction guide", "product support", "warning", "batteries only"]:
                    if phrase in ocr_lower:
                        matched_features.append(phrase.title())
                        
                if matched_features:
                    mismatches.append({
                        "status": "PASS",
                        "photo_num": photo_count,
                        "message": f"3F/FFF 照片 {photo_count} 辨識內容與圖面規範相符（辨識特徵: {', '.join(matched_features)}）"
                    })
                else:
                    mismatches.append({
                        "status": "INFO",
                        "photo_num": photo_count,
                        "message": f"3F/FFF 照片 {photo_count} 辨識文字: '{ocr_text[:60]}...'"
                    })
        except Exception as e:
            mismatches.append({
                "status": "ERROR",
                "message": f"處理 3F/FFF 照片 {idx+1} 時發生錯誤: {e}"
            })
            
    return table_findings + mismatches


def run_embedded_fai_discrepancy_audit(pqp_path, wb):
    """
    Audits embedded raw FAI workbooks (e.g. xl/embeddings/Microsoft_Excel_Worksheet*.xlsx)
    attached to 100% Dimensional sheets and flags:
    1. Non-numeric / visual FAIL items (e.g. Balloon 37 Paint Masking Area 'FAIL' or 'NG').
    2. Contradictory / Concealed FAIL items where Measurement says 'FAIL' but Results says 'O' / 'OK'.
    3. Balloon Omission / Table Truncation where embedded raw FAI has balloons (e.g. #32~#37) not shown in outer PQP sheet.
    4. Sibling embedded workbook duplication (Cavity 1 duplicated into Cavity 2).
    5. Any sample measurement in raw FAI that is Out-of-Spec (OOS, % Tol Used > 100%).
    6. Discrepancies between the raw FAI measurement and the active PQP 100% Dimensional sheet, paired accurately via get_embedded_to_sheets_map.
    """
    import zipfile, io, os, re
    results = []
    seen_discrepancies = set()
    
    try:
        with zipfile.ZipFile(io.BytesIO(read_file_bytes_safe(pqp_path)), 'r') as z:
            emb_to_sheets = get_embedded_to_sheets_map(z)
            pqp_dim_sheets = [s for s in wb.sheetnames if "dimensional" in s.lower() and wb[s].sheet_state == 'visible']
            
            embedded_fai_files = [n for n in z.namelist() if n.startswith('xl/embeddings/') and (n.endswith('.xlsx') or n.endswith('.xls'))]
            
            emb_signatures = {}  # emb_base -> {(balloon, sample): meas_str}
            emb_host_map = {}    # emb_base -> host_sheet
            
            for emb_name in embedded_fai_files:
                emb_base = os.path.basename(emb_name)
                emb_info = z.getinfo(emb_name)
                if emb_info.file_size > 10 * 1024 * 1024:
                    continue
                    
                data = z.read(emb_name)
                
                # Pre-check sheet names using zipfile before loading full workbook
                fai_relevant = False
                try:
                    with zipfile.ZipFile(io.BytesIO(data), 'r') as emb_z:
                        if 'xl/workbook.xml' in emb_z.namelist():
                            import xml.etree.ElementTree as ET
                            wb_xml = emb_z.read('xl/workbook.xml')
                            wb_tree = ET.fromstring(wb_xml)
                            ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                            for sheet in wb_tree.findall(f'{{{ns}}}sheets/{{{ns}}}sheet'):
                                sn = sheet.attrib.get('name', '').lower()
                                if any(k in sn for k in ['dim', '100%', 'fai', 'result']):
                                    fai_relevant = True
                                    break
                except Exception:
                    pass
                    
                if not fai_relevant:
                    continue
                    
                emb_wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
                reported_keys = set()
                
                # Check for dimensional sheets in embedded workbook
                for sname in emb_wb.sheetnames:
                    if any(k in sname.lower() for k in ['dim', '100%', 'fai', 'result']):
                        ws_emb = emb_wb[sname]

                        # Check for formula error codes in embedded FAI sheet header or cells
                        for er_r in range(1, min(20, ws_emb.max_row + 1)):
                            for er_c in range(1, min(20, ws_emb.max_column + 1)):
                                ev = str(ws_emb.cell(row=er_r, column=er_c).value or '').strip()
                                if ev in ["#VALUE!", "#REF!", "#DIV/0!", "#NAME?"]:
                                    err_key = (emb_name, sname, er_r, er_c)
                                    if err_key not in reported_keys:
                                        reported_keys.add(err_key)
                                        results.append({
                                            "status": "WARNING",
                                            "sheet": "100% Dimensional Results",
                                            "embedded_file": emb_base,
                                            "message": f"嵌入的原始 FAI 活頁簿 [{emb_base}] 分頁 [{sname}] 於 Row {er_r}, Col {er_c} 存在公式錯誤碼 '{ev}'。"
                                        })

                        # Check for Revision header field in embedded sheet
                        rev_found = False
                        rev_has_val = False
                        for er_r in range(1, min(15, ws_emb.max_row + 1)):
                            for er_c in range(1, min(20, ws_emb.max_column + 1)):
                                lbl = str(ws_emb.cell(row=er_r, column=er_c).value or '').strip().lower()
                                if lbl in ["revision", "rev", "rev."]:
                                    rev_found = True
                                    for er_co in range(1, 6):
                                        nv = ws_emb.cell(row=er_r, column=er_c+er_co).value
                                        if nv is not None and str(nv).strip() != "":
                                            rev_has_val = True
                                            break
                        if rev_found and not rev_has_val:
                            rev_key = (emb_name, sname, "REV_BLANK")
                            if rev_key not in reported_keys:
                                reported_keys.add(rev_key)
                                results.append({
                                    "status": "WARNING",
                                    "sheet": "100% Dimensional Results",
                                    "embedded_file": emb_base,
                                    "message": f"嵌入的原始 FAI 活頁簿 [{emb_base}] 分頁 [{sname}] 表頭 Revision 欄位為空白，未填寫版次。"
                                        })

                        # Dynamically determine columns for ws_emb
                        e_sample_col = 1
                        e_balloon_col = 2
                        e_nom_col = 3
                        e_usl_col = 6
                        e_lsl_col = 7
                        e_meas_col = 8
                        e_tol_col = 9
                        e_ctq_col = 12
                        e_res_col = 13
                        
                        for hr in range(18, min(27, ws_emb.max_row + 1)):
                            row_txt = [str(ws_emb.cell(row=hr, column=hc).value or '').strip().lower() for hc in range(1, min(20, ws_emb.max_column + 1))]
                            if any('balloon' in t or '尺寸#' in t or '球標' in t for t in row_txt):
                                for hc, t in enumerate(row_txt, 1):
                                    if 'sample' in t or '樣品#' in t: e_sample_col = hc
                                    elif 'balloon' in t or '尺寸#' in t or '球標' in t: e_balloon_col = hc
                                    elif 'ctq' in t or '重點尺寸' in t: e_ctq_col = hc
                                    elif 'nominal' in t or '標準值' in t: e_nom_col = hc
                                    elif 'upper\nspec' in t or 'upper spec' in t: e_usl_col = hc
                                    elif 'lower\nspec' in t or 'lower spec' in t: e_lsl_col = hc
                                    elif ('meas' in t or '量測值' in t) and 'method' not in t and 'tool' not in t: e_meas_col = hc
                                    elif ('tol' in t or '上限%' in t) and 'upper' not in t: e_tol_col = hc
                                    elif 'result' in t or 'conformance' in t or '判定' in t or '結果' in t: e_res_col = hc
                                break

                        emb_rows = []
                        emb_meas_dict = {}
                        max_scan_row = min(ws_emb.max_row, 2000)
                        for r in range(1, max_scan_row + 1):
                            s_val = ws_emb.cell(row=r, column=e_sample_col).value
                            b_val = ws_emb.cell(row=r, column=e_balloon_col).value
                            if s_val is None or b_val is None:
                                continue
                            try:
                                s_int = int(float(str(s_val).strip()))
                                b_int = int(float(str(b_val).strip()))
                            except Exception:
                                continue

                            nom_val = ws_emb.cell(row=r, column=e_nom_col).value
                            usl_val = ws_emb.cell(row=r, column=e_usl_col).value
                            lsl_val = ws_emb.cell(row=r, column=e_lsl_col).value
                            meas_val = ws_emb.cell(row=r, column=e_meas_col).value
                            tol_used_val = ws_emb.cell(row=r, column=e_tol_col).value
                            ctq_val = str(ws_emb.cell(row=r, column=e_ctq_col).value or '').strip()
                            res_val = str(ws_emb.cell(row=r, column=e_res_col).value or '').strip()

                            meas_str = str(meas_val or '').strip()
                            meas_f = None
                            try:
                                meas_f = float(meas_val)
                            except Exception:
                                pass

                            tol_used_f = None
                            try:
                                tol_used_f = float(tol_used_val)
                            except Exception:
                                pass

                            emb_rows.append({
                                "sample": s_int,
                                "balloon": b_int,
                                "nominal": nom_val,
                                "usl": usl_val,
                                "lsl": lsl_val,
                                "measurement": meas_f,
                                "meas_raw": meas_val,
                                "meas_str": meas_str,
                                "tol_used": tol_used_f,
                                "ctq": ctq_val,
                                "result": res_val,
                                "row": r
                            })
                            emb_meas_dict[(b_int, s_int)] = meas_str

                        if not emb_rows:
                            continue
                            
                        emb_signatures[emb_base] = emb_meas_dict

                        # Smart pairing: prioritize exact host worksheet from embedded relationships
                        host_sheets = emb_to_sheets.get(emb_base, [])
                        matched_pqp = [s for s in host_sheets if s in pqp_dim_sheets]
                        if matched_pqp:
                            target_pqp_sheets = matched_pqp
                        else:
                            target_pqp_sheets = []
                            m_pns = re.findall(r'(\d{5}(?:-\d{3})?)', sname)
                            for ps in pqp_dim_sheets:
                                if m_pns and any(p in ps for p in m_pns):
                                    target_pqp_sheets.append(ps)
                            if not target_pqp_sheets:
                                sname_clean = re.sub(r'[^a-zA-Z0-9]', '', sname).lower()
                                for ps in pqp_dim_sheets:
                                    ps_clean = re.sub(r'[^a-zA-Z0-9]', '', ps).lower()
                                    if sname_clean in ps_clean or ps_clean in sname_clean:
                                        target_pqp_sheets.append(ps)
                            if not target_pqp_sheets:
                                target_pqp_sheets = pqp_dim_sheets

                        emb_host_map[emb_base] = target_pqp_sheets[0] if target_pqp_sheets else ""

                        # Check 1: Non-numeric / visual FAIL or Contradictory Pass in raw FAI
                        for er in emb_rows:
                            b_id = er["balloon"]
                            s_id = er["sample"]
                            m_str = er["meas_str"]
                            r_str = er["result"]
                            nom_desc = er["nominal"] or f"Balloon {b_id}"
                            
                            is_meas_fail = m_str.upper() in ['FAIL', 'NG', 'REJECT', 'REJ', '不合格']
                            is_res_fail = r_str.upper() in ['FAIL', 'NG', 'REJECT', 'REJ', 'X', '不合格']
                            
                            if is_meas_fail or is_res_fail:
                                fail_key = (emb_base, b_id, s_id, "FAIL_ITEM")
                                if fail_key not in seen_discrepancies:
                                    seen_discrepancies.add(fail_key)
                                    h_sheet = target_pqp_sheets[0] if target_pqp_sheets else "100% Dimensional Results"
                                    if is_meas_fail and r_str.upper() in ['O', 'OK', 'PASS', '合格', '符合']:
                                        results.append({
                                            "status": "FAIL",
                                            "sheet": h_sheet,
                                            "embedded_file": emb_base,
                                            "balloon": b_id,
                                            "sample": s_id,
                                            "message": f"❌ **FAI 判定矛盾與隱匿 (Contradictory / Concealed FAIL)**: 嵌入的原始 FAI 活頁簿 [{emb_base}] 中 Balloon {b_id} (Sample #{s_id}, 項目: '{nom_desc}') 量測值明確記載為 '{m_str}' (FAIL)，但判定欄卻填寫 '{r_str}' (OK)，且該項目未在外層 PQP 尺寸表中如實揭露！"
                                        })
                                    else:
                                        results.append({
                                            "status": "FAIL",
                                            "sheet": h_sheet,
                                            "embedded_file": emb_base,
                                            "balloon": b_id,
                                            "sample": s_id,
                                            "message": f"❌ **原始 FAI 檢驗不合格 (FAIL Item)**: 嵌入的原始 FAI 活頁簿 [{emb_base}] 中 Balloon {b_id} (Sample #{s_id}, 項目: '{nom_desc}') 判定為 '{m_str or r_str}' (FAIL / 不合格)！"
                                        })

                        # Check 2: Outer vs Embedded comparison
                        for p_sname in target_pqp_sheets:
                            p_ws = wb[p_sname]
                            p_cols = find_dimensional_columns(p_ws)
                            
                            pqp_entries = {}
                            pqp_balloons = set()
                            for pr in range(25, p_ws.max_row + 1):
                                pb = p_ws.cell(row=pr, column=p_cols["balloon"]).value
                                ps = p_ws.cell(row=pr, column=p_cols["sample"]).value
                                pm = p_ws.cell(row=pr, column=p_cols["meas"]).value
                                pctq = str(p_ws.cell(row=pr, column=p_cols.get("ctq", 3)).value or '').strip()
                                if pb is not None and ps is not None:
                                    try:
                                        b_num = int(float(str(pb).strip()))
                                        s_num = int(float(str(ps).strip()))
                                        pqp_balloons.add(b_num)
                                        p_nom_val = p_ws.cell(row=pr, column=p_cols["nominal"]).value
                                        pqp_entries[(b_num, s_num)] = {
                                            "meas": float(pm) if pm is not None and str(pm).strip() not in ['', 'None'] else None,
                                            "meas_raw": pm,
                                            "nom": p_nom_val,
                                            "ctq": pctq,
                                            "row": pr
                                        }
                                    except (ValueError, TypeError):
                                        pass
                                        
                            # Check 2.1: Balloon Omission / Table Truncation, Re-ballooned Legacy FAI & Abbreviated FAI Diagnostic
                            emb_balloons = set(er["balloon"] for er in emb_rows)
                            missing_in_pqp = emb_balloons - pqp_balloons

                            # Detect if embedded FAI uses a completely different / legacy balloon numbering order
                            common_balloons = sorted(list(emb_balloons & pqp_balloons))
                            nom_mismatch_cnt = 0
                            nom_compare_cnt = 0
                            for cb in common_balloons:
                                er_cb = next((x for x in emb_rows if x["balloon"] == cb), None)
                                pe_cb = pqp_entries.get((cb, 1)) or next((v for (k_b, _), v in pqp_entries.items() if k_b == cb), None)
                                if er_cb and pe_cb and er_cb.get("nominal") is not None and pe_cb.get("nom") is not None:
                                    try:
                                        en_f = float(er_cb["nominal"])
                                        pn_f = float(pe_cb["nom"])
                                        nom_compare_cnt += 1
                                        if abs(en_f - pn_f) > 0.01:
                                            nom_mismatch_cnt += 1
                                    except Exception:
                                        pass
                            is_reballooned_legacy = (nom_compare_cnt >= 5 and (nom_mismatch_cnt / nom_compare_cnt) >= 0.5 and len(emb_balloons) > 10)

                            if is_reballooned_legacy:
                                rebal_key = (emb_base, p_sname, "REBALLOONED_LEGACY_FAI")
                                if rebal_key not in seen_discrepancies:
                                    seen_discrepancies.add(rebal_key)
                                    results.append({
                                        "status": "WARNING",
                                        "sheet": p_sname,
                                        "embedded_file": emb_base,
                                        "message": f"⚠️ **內嵌原始 FAI 活頁簿球標版次未同步 (Legacy / Re-ballooned FAI Attachment)**: 嵌入的原始 FAI 活頁簿 [{emb_base}] 仍為舊版 {len(emb_balloons)} 項球標編號體系（#1~#{max(emb_balloons)}），而外層 PQP 工作表 [{p_sname}] 已依新版圖面重新編號為 {len(pqp_balloons)} 項球標體系（#1~#{max(pqp_balloons)}，兩者同球標號之公稱值順序不同）。建議送審前將內嵌檔案更新為與外層一致的 {len(pqp_balloons)} 項球標版 FAI 報表。"
                                    })
                                continue

                            # Diagnostic: Supplier abbreviated key-characteristics FAI pattern
                            if len(emb_balloons) <= 10 and len(pqp_balloons) >= 20:
                                abbrev_key = (emb_base, p_sname, "ABBREV_FAI")
                                if abbrev_key not in seen_discrepancies:
                                    seen_discrepancies.add(abbrev_key)
                                    has_ref_note = False
                                    for r_idx in range(1, min(25, ws_emb.max_row + 1)):
                                        for c_idx in range(1, min(10, ws_emb.max_column + 1)):
                                            c_val = str(ws_emb.cell(row=r_idx, column=c_idx).value or '')
                                            if any(kw in c_val for kw in ['附檔', '重點尺寸', 'attached', 'key dim', 'Titan Project']):
                                                has_ref_note = True
                                                break
                                    results.append({
                                        "status": "INFO",
                                        "sheet": p_sname,
                                        "embedded_file": emb_base,
                                        "message": f"ℹ️ **供應商簡約 FAI 模式識別**: 內嵌 FAI 活頁簿 [{emb_base}] 僅記載 {len(emb_balloons)} 項重點檢驗（如 SC 肉厚與外觀/材質規範），外層 PQP 則已完整展開 {len(pqp_balloons)} 項全尺寸" + ("，且 FAI 備註已註明依重點尺寸量測" if has_ref_note else "") + "。建議送審前將外層完整全尺寸量測數據同步覆蓋回內嵌活頁簿以維持完整性。"
                                    })

                            if missing_in_pqp and pqp_balloons:
                                omiss_key = (p_sname, emb_base, "OMISSION")
                                if omiss_key not in seen_discrepancies:
                                    seen_discrepancies.add(omiss_key)
                                    omitted_details = []
                                    for mb in sorted(list(missing_in_pqp)):
                                        mb_rows = [er for er in emb_rows if er["balloon"] == mb]
                                        nom_str = mb_rows[0]["nominal"] if mb_rows else ""
                                        m_val_str = mb_rows[0]["meas_str"] if mb_rows else ""
                                        omitted_details.append(f"#{mb} ({nom_str}: {m_val_str})")
                                    results.append({
                                        "status": "FAIL",
                                        "sheet": p_sname,
                                        "embedded_file": emb_base,
                                        "message": f"❌ **外層 PQP 尺寸清單截斷漏列 (Balloon Omission)**: 嵌入的原始 FAI 包含 {len(missing_in_pqp)} 個後續檢驗項目 {', '.join(omitted_details)}，但外層 PQP 工作表 [{p_sname}] 僅列至球標 #{max(pqp_balloons)}，後續項目遭到截斷遺漏！"
                                    })

                            # Check 2.2: Numeric measurements & Copy-Paste Shift Detection
                            for er in emb_rows:
                                b_id = er["balloon"]
                                s_id = er["sample"]
                                emb_meas = er["measurement"]
                                emb_tol = er["tol_used"]
                                emb_nom = er.get("nominal")
                                
                                # Out of spec in raw FAI
                                if emb_tol is not None and (emb_tol > 1.0 or emb_tol < -1.0):
                                    oos_key = (p_sname, b_id, s_id, "OOS")
                                    if oos_key not in seen_discrepancies:
                                        seen_discrepancies.add(oos_key)
                                        results.append({
                                            "status": "ALERT",
                                            "sheet": p_sname,
                                            "embedded_file": emb_base,
                                            "balloon": b_id,
                                            "sample": s_id,
                                            "emb_measurement": emb_meas,
                                            "tol_used_pct": f"{emb_tol*100:.1f}%",
                                            "message": f"嵌入的原始 FAI 活頁簿 [{emb_base}] 中 Balloon {b_id} (Sample #{s_id}) 量測值為 {emb_meas}，公差耗用達 {emb_tol*100:.1f}%（超差 Out-of-Spec）！"
                                        })
                                    
                                # Discrepancy between raw FAI and PQP sheet
                                if (b_id, s_id) in pqp_entries:
                                    p_entry = pqp_entries[(b_id, s_id)]
                                    p_meas = p_entry["meas"]
                                    p_nom = p_entry.get("nom")

                                    # Check nominal alignment: if nominals completely differ, this is a balloon index remapping
                                    nom_align_ok = True
                                    if emb_nom is not None and p_nom is not None:
                                        try:
                                            if abs(float(emb_nom) - float(p_nom)) > 0.01:
                                                nom_align_ok = False
                                        except Exception:
                                            pass

                                    if not nom_align_ok:
                                        remap_key = (emb_base, b_id, "BALLOON_REMAP")
                                        if remap_key not in seen_discrepancies:
                                            seen_discrepancies.add(remap_key)
                                            target_pqp_b = None
                                            for pb_k, pe_v in pqp_entries.items():
                                                if pe_v.get("nom") is not None:
                                                    try:
                                                        if abs(float(pe_v["nom"]) - float(emb_nom)) < 0.005:
                                                            target_pqp_b = pb_k[0]
                                                            break
                                                    except Exception:
                                                        pass
                                            if target_pqp_b is not None:
                                                results.append({
                                                    "status": "INFO",
                                                    "sheet": p_sname,
                                                    "embedded_file": emb_base,
                                                    "balloon": b_id,
                                                    "message": f"ℹ️ **FAI 球標重對應識別**: 內嵌 FAI [{emb_base}] 之 Balloon #{b_id} (Nominal: {emb_nom}) 與外層 PQP Balloon #{b_id} (Nominal: {p_nom}) 公稱值不同，經比對實對應至外層 PQP 之 Balloon #{target_pqp_b} (Nominal: {emb_nom})。此為供應商簡約首件編號與全尺寸球標重新編號對應，非量測數據竄改。"
                                                })
                                            else:
                                                results.append({
                                                    "status": "WARNING",
                                                    "sheet": p_sname,
                                                    "embedded_file": emb_base,
                                                    "balloon": b_id,
                                                    "message": f"內嵌 FAI [{emb_base}] 之 Balloon #{b_id} 公稱值 ({emb_nom}) 與外層 PQP 同球標公稱值 ({p_nom}) 存在顯著差異。"
                                                })
                                    else:
                                        if p_meas is not None and emb_meas is not None:
                                            if abs(p_meas - emb_meas) > 0.0001:
                                                copied_from_sheet = None
                                                for other_sname in emb_wb.sheetnames:
                                                    if other_sname != sname and any(k in other_sname.lower() for k in ['dim', '100%', 'fai', 'result']):
                                                        other_ws = emb_wb[other_sname]
                                                        for or_idx in range(15, min(other_ws.max_row + 1, 100)):
                                                            o_s = other_ws.cell(row=or_idx, column=e_sample_col).value
                                                            o_b = other_ws.cell(row=or_idx, column=e_balloon_col).value
                                                            o_m = other_ws.cell(row=or_idx, column=e_meas_col).value
                                                            try:
                                                              if (
                                                                  int(
                                                                      float(
                                                                          str(
                                                                              o_b
                                                                          ).strip()
                                                                      )
                                                                  )
                                                                  == b_id
                                                                  and int(
                                                                      float(
                                                                          str(
                                                                              o_s
                                                                          ).strip()
                                                                      )
                                                                  )
                                                                  == s_id
                                                                  and o_m
                                                                  is not None
                                                              ):
                                                                if (
                                                                    abs(
                                                                        float(
                                                                            o_m
                                                                        )
                                                                        - p_meas
                                                                    )
                                                                    < 0.0001
                                                                ):
                                                                  copied_from_sheet = (
                                                                      other_sname
                                                                  )
                                                                  break
                                                            except Exception:
                                                              pass
                                                        if copied_from_sheet:
                                                            break
                                                
                                                disc_key = (p_sname, b_id, s_id, round(p_meas, 5), round(emb_meas, 5))
                                                if disc_key not in seen_discrepancies:
                                                    seen_discrepancies.add(disc_key)
                                                    if copied_from_sheet:
                                                        results.append({
                                                            "status": "FAIL",
                                                            "sheet": p_sname,
                                                            "embedded_file": emb_base,
                                                            "balloon": b_id,
                                                            "sample": s_id,
                                                            "emb_measurement": emb_meas,
                                                            "pqp_measurement": p_meas,
                                                            "message": f"❌ **跨分頁數據錯位複製 (Copy-Paste Shift)**: 分頁 [{p_sname}] 中 Balloon {b_id} (Sample #{s_id}) 填寫之量測值 {p_meas} 與本料號原始 FAI 數據 ({emb_meas}) 不符！經深度比對，此數值與隔壁料號分頁 [{copied_from_sheet}] 的量測值完全一致，疑似由相鄰分頁錯位複製貼上！"
                                                        })
                                                    else:
                                                        results.append({
                                                            "status": "WARNING",
                                                            "sheet": p_sname,
                                                            "embedded_file": emb_base,
                                                            "balloon": b_id,
                                                            "sample": s_id,
                                                            "emb_measurement": emb_meas,
                                                            "pqp_measurement": p_meas,
                                                            "message": f"嵌入的原始 FAI 活頁簿 [{emb_base}] (Balloon {b_id}, Sample #{s_id}) 量測值為 {emb_meas}，但外層 PQP 工作表 [{p_sname}] 中數值為 {p_meas}（數據不一致）！"
                                                        })
                                            
                                    # Check CTQ/SC designation consistency
                                    p_ctq = pqp_entries[(b_id, s_id)]["ctq"]
                                    if p_ctq in ["SC", "CTQ"] and er["ctq"] not in ["SC", "CTQ", "V", "O", "1"]:
                                        ctq_key = (p_sname, b_id, "CTQ_MISMATCH")
                                        if ctq_key not in seen_discrepancies:
                                            seen_discrepancies.add(ctq_key)
                                            results.append({
                                                "status": "WARNING",
                                                "sheet": p_sname,
                                                "embedded_file": emb_base,
                                                "balloon": b_id,
                                                "message": f"主 PQP 分頁 [{p_sname}] 中 Balloon {b_id} 宣告為 '{p_ctq}'，但在嵌入 FAI [{emb_base}] 分頁 [{sname}] 中此欄位為空白未標註！"
                                            })

            # Check 3: Sibling Embedded Workbook Duplication (Cavity Mismatch)
            emb_bases = list(emb_signatures.keys())
            for i in range(len(emb_bases)):
                for j in range(i + 1, len(emb_bases)):
                    f1, f2 = emb_bases[i], emb_bases[j]
                    h1, h2 = emb_host_map.get(f1, ""), emb_host_map.get(f2, "")
                    if h1 != h2 and h1 in pqp_dim_sheets and h2 in pqp_dim_sheets:
                        sig1, sig2 = emb_signatures[f1], emb_signatures[f2]
                        common_keys = set(sig1.keys()) & set(sig2.keys())
                        if len(common_keys) >= 20:
                            matches = sum(1 for k in common_keys if sig1[k] == sig2[k])
                            if matches == len(common_keys):
                                dup_key = (f1, f2, "DUP_EMBEDDING")
                                if dup_key not in seen_discrepancies:
                                    seen_discrepancies.add(dup_key)
                                    results.append({
                                        "status": "FAIL",
                                        "sheet": h2,
                                        "embedded_file": f2,
                                        "message": f"❌ **嵌入附件重複/穴號未更換 (Duplicate Cavity Embedding)**: 工作表 [{h2}] 內嵌之原始 FAI 活頁簿 [{f2}] 數據與 [{h1}] 之 [{f1}] 100% 完全相同！外層表格分別填寫了不同穴號數據，但內嵌檔案疑似因複製貼上而漏更換為第 2 穴之原始記錄！"
                                    })
    except Exception:
        pass
        
    return results


def get_embedded_to_sheets_map(z):
    """
    Returns a dict mapping embedded object basename (e.g. 'oleObject1.bin') to a list of sheet names where it is embedded.
    """
    try:
        wb_tree = ET.fromstring(z.read('xl/workbook.xml'))
        wb_rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rel_to_target = {rel.attrib['Id']: rel.attrib['Target'] for rel in wb_rels}
        
        sheet_map = {}
        for s in wb_tree.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
            s_name = s.attrib['name']
            r_id = s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            if r_id in rel_to_target:
                target = rel_to_target[r_id].lstrip('/')
                if not target.startswith('worksheets/'):
                    target = 'worksheets/' + os.path.basename(target)
                sheet_map[target] = s_name

        drawing_to_sheet = {}
        for file in z.namelist():
            if file.startswith('xl/worksheets/_rels/') and file.endswith('.rels'):
                sheet_target = 'worksheets/' + file.split('/')[-1][:-5]
                sheet_name = sheet_map.get(sheet_target, sheet_target)
                try:
                    rels_tree = ET.fromstring(z.read(file))
                    for rel in rels_tree:
                        target = rel.attrib.get('Target', '')
                        if 'drawing' in target.lower():
                            d_norm = 'drawings/' + os.path.basename(target)
                            drawing_to_sheet[d_norm] = sheet_name
                except Exception:
                    pass

        obj_to_sheets = {}
        # 1. Check worksheet rels
        for file in z.namelist():
            if file.startswith('xl/worksheets/_rels/') and file.endswith('.rels'):
                sheet_target = 'worksheets/' + file.split('/')[-1][:-5]
                sheet_name = sheet_map.get(sheet_target, sheet_target)
                try:
                    rels_tree = ET.fromstring(z.read(file))
                    for rel in rels_tree:
                        target = rel.attrib.get('Target', '')
                        if 'embeddings' in target.lower() or 'oleobject' in target.lower():
                            obj_name = os.path.basename(target)
                            if sheet_name not in obj_to_sheets.setdefault(obj_name, []):
                                obj_to_sheets[obj_name].append(sheet_name)
                except Exception:
                    pass

        # 2. Check drawing rels
        for file in z.namelist():
            if file.startswith('xl/drawings/_rels/') and file.endswith('.rels'):
                drawing_target = 'drawings/' + file.split('/')[-1][:-5]
                sheet_name = drawing_to_sheet.get(drawing_target, drawing_to_sheet.get('drawings/' + os.path.basename(drawing_target), 'Drawing'))
                try:
                    rels_tree = ET.fromstring(z.read(file))
                    for rel in rels_tree:
                        target = rel.attrib.get('Target', '')
                        if 'embeddings' in target.lower() or 'oleobject' in target.lower():
                            obj_name = os.path.basename(target)
                            if sheet_name not in obj_to_sheets.setdefault(obj_name, []):
                                obj_to_sheets[obj_name].append(sheet_name)
                except Exception:
                    pass

        return obj_to_sheets
    except Exception:
        return {}


def extract_expected_description(wb=None, dwg_spec_data=None):
    # Check PQP workbook first as it contains clean, dedicated metadata cells
    if wb:
        mapping_desc = ["part description", "description", "part desc", "品名", "產品名稱"]
        search_order = ['TitlePg', 'PQS', 'Control Plan', '100% Dimensional Results', 'Inspection Plan']
        sheets = [wb[s] for s in search_order if s in wb.sheetnames] + [s for s in wb.worksheets if s.title not in search_order]
        for ws in sheets:
            if ws.sheet_state != 'visible':
                continue
            for r in range(1, min(25, ws.max_row + 1)):
                for c in range(1, min(25, ws.max_column + 1)):
                    v = str(ws.cell(row=r, column=c).value or '').strip()
                    v_lower = v.lower()
                    if any(v_lower == t or v_lower.startswith(t + ":") or v_lower.startswith(t + " ") for t in mapping_desc):
                        val = None
                        if ":" in v:
                            val = v.split(":", 1)[1].strip()
                        if not val:
                            for offset in range(1, 15):
                                pv = ws.cell(row=r, column=c + offset).value
                                if pv is not None and str(pv).strip():
                                    pv_str = str(pv).strip()
                                    if not any(pv_str.lower() == t for t in mapping_desc):
                                        val = pv_str
                                        break
                        if val and len(val) >= 3 and not any(k in val.upper() for k in ["DRAWING NUMBER", "REVISION", "DWG NO"]):
                            return val

    # Fallback to dwg_spec_data if available
    if dwg_spec_data and isinstance(dwg_spec_data, dict):
        m = dwg_spec_data.get("metadata", {})
        if m.get("title") and str(m["title"]).strip().upper() not in ["SIZE", "DWG NO", "REV", "SCALE", "SHEET"]:
            return str(m["title"]).strip()
        dwg_text = dwg_spec_data.get("text", "")
        m_dwg = re.search(r'\bTITLE\b[\s\S]{1,120}?(INSERT[^\r\n]+|DIVIDER[^\r\n]+|[A-Z0-9\s,-]{5,40})(?:\s+SIZE|\s+C|\s+D|\s+DO NOT SCALE|\s+DWG|\s+REV)', dwg_text)
        if m_dwg:
            cand = m_dwg.group(1).strip()
            if not any(k in cand.upper() for k in ["SIZE", "DWG NO", "SCALE", "SHEET", "DRAFTER", "CHECKED"]):
                return cand

    return None


def run_embedded_attachments_audit(pqp_path, wb=None, dwg_spec_data=None):
    """
    Scans embedded OLE objects (.bin) and verifies attached certificates and test reports with their hosting worksheets.
    Cross-checks Product Name / Description against Drawing Title and PQP Description.
    """
    import zipfile
    reports = []
    expected_desc = extract_expected_description(wb, dwg_spec_data) if (wb or dwg_spec_data) else None
    
    try:
        with zipfile.ZipFile(io.BytesIO(read_file_bytes_safe(pqp_path)), 'r') as z:
            obj_to_sheets = get_embedded_to_sheets_map(z)
            for name in z.namelist():
                if name.startswith('xl/embeddings/') and name.endswith('.bin'):
                    data = z.read(name)
                    pdf_idx = data.find(b'%PDF')
                    if pdf_idx != -1:
                        pdf_bytes = data[pdf_idx:]
                        try:
                            import pymupdf
                            doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
                            num_pages = len(doc)
                            first_txt = (doc[0].get_text() or "").replace('\n', ' ') if num_pages > 0 else ""
                            all_txt = " ".join([p.get_text() or "" for p in doc[:3]])
                            
                            rep_type = "未定義文件"
                            obj_base = os.path.basename(name)
                            sheets_found = obj_to_sheets.get(obj_base, [])
                            is_dwg_sheet = any('drawing' in s.lower() for s in sheets_found) if sheets_found else False
                            is_finish_sheet = any('finish' in s.lower() for s in sheets_found) if sheets_found else False
                            
                            if is_dwg_sheet or any(k in all_txt.upper() for k in ["DO NOT SCALE", "THIRD ANGLE PROJECTION", "DWG NO", "DRAFTER"]):
                                rep_type = "工程圖面 (Engineering Drawing PDF)"
                            elif is_finish_sheet or any(k in all_txt.upper() for k in ["SALT SPRAY", "ASTM B117", "FINISH TEST REPORT", "NSS (NEUTRAL SALT SPRAY)"]):
                                rep_type = "鹽霧與表面處理測試報告 (Finish / Salt Spray Test Report)"
                            elif "ROHS" in all_txt.upper() or "SGS" in all_txt.upper() or "VOC" in all_txt.upper() or "REACH" in all_txt.upper():
                                rep_type = "SGS RoHS / 綠色環保認證報告"
                            elif "ISTA" in all_txt.upper() or "DROP TEST" in all_txt.upper() or ("PACKAGING" in all_txt.upper() and ("TEST" in all_txt.upper() or "DROP" in all_txt.upper())):
                                rep_type = "ISTA 3A 包裝落下測試報告"
                            elif "SPECIFICATION" in all_txt.upper() or "說明書" in all_txt or "產品說明書" in all_txt:
                                rep_type = "材料規格說明書 (Material Specification)"
                            elif "DRAWING" in all_txt.upper():
                                rep_type = "工程圖面 (Engineering Drawing PDF)"
                            sheet_str = ', '.join([f'[{s}]' for s in sheets_found]) if sheets_found else '未指定分頁'
                            
                            # Check Product Name / Description against Drawing & PQP
                            desc_found = None
                            desc_status = "none"
                            desc_msg = ""
                            IGNORE_STATUS_WORDS = {
                                "RELEASED", "DRAFT", "PENDING", "APPROVED", "COMPLETED", "IN PROGRESS",
                                "N/A", "NONE", "TEST PROTOCOL", "SN ID", "SN", "ID", "SAMPLE ID", "ITEM NO",
                                "MANAGER", "SPECIALIST", "JUNIOR SPECIALIST", "SENIOR SPECIALIST", "ENGINEER",
                                "SUPERVISOR", "DIRECTOR", "PROJECT MGR", "DEPARTMENT", "PREPARED BY", "APPROVED BY"
                            }
                            # Prioritize Sample Name / 样品名称 / Sample Description from certificate header
                            for m_desc in re.finditer(r'(?:Sample\s+Name|样品名称|Samples?\s+Description|Product\s+Name|Part\s+Description|Item\s+Name|品名|產品名稱|品名規格)\s*[:：\n]\s*([^\r\n]+)', all_txt, re.IGNORECASE):
                                raw_desc = m_desc.group(1).strip()
                                desc_cand = re.sub(r'\b(?:Part\s+Number|Material|Rev|Date)\b.*$', '', raw_desc, flags=re.IGNORECASE).strip()
                                desc_cand = re.sub(r'^(?:SN\s*ID|SAMPLE\s*ID)\b.*$', '', desc_cand, flags=re.IGNORECASE).strip()
                                if desc_cand and desc_cand.upper() not in IGNORE_STATUS_WORDS and len(desc_cand) >= 2:
                                    desc_found = desc_cand
                                    break
                            if desc_found:
                                is_mat_cert = any('material' in s.lower() for s in sheets_found)
                                if is_mat_cert:
                                    desc_status = "raw_material"
                                    desc_msg = f"報告記載材料品名為 '{desc_found}'，屬於 [Material Certs] 原材料規格/RoHS證明"
                                elif expected_desc and desc_found:
                                    tokens_found = set(re.findall(r'[A-Za-z0-9\u4e00-\u9fff]{2,}', desc_found.upper()))
                                    tokens_exp = set(re.findall(r'[A-Za-z0-9\u4e00-\u9fff]{2,}', expected_desc.upper()))
                                    stopwords = {"THE", "AND", "FOR", "OF", "WITH", "LTD", "CO", "INC", "CORP", "UNIT", "KS"}
                                    tokens_found -= stopwords
                                    tokens_exp -= stopwords
                                    # Also check full text of Finish Test Reports if header cell shifted due to multi-column layout
                                    txt_upper = all_txt.upper().replace("BLACK PLATE", "BACK PLATE")
                                    exp_in_body = sum(1 for t in tokens_exp if t in txt_upper) >= max(1, len(tokens_exp) // 2)
                                    if tokens_found and tokens_exp and not (tokens_found & tokens_exp) and not exp_in_body:
                                        desc_status = "mismatch"
                                        desc_msg = f"報告記載品名為 '{desc_found}'，但圖面/PQP品名為 '{expected_desc}'（疑似複製外包裝或其他零件之 CoC 範本未改品名）"
                                    elif (tokens_found & tokens_exp) or exp_in_body:
                                        desc_status = "match"
                                        desc_msg = f"報告記載品名與圖面/PQP品名 ('{expected_desc}') 相符"

                            reports.append({
                                "file": obj_base,
                                "sheets": sheets_found,
                                "sheet_str": sheet_str,
                                "pages": num_pages,
                                "type": rep_type,
                                "summary": first_txt[:80] + "...",
                                "desc_found": desc_found,
                                "desc_status": desc_status,
                                "desc_msg": desc_msg
                            })
                        except Exception:
                            pass
    except Exception:
        pass
        
    return reports


def get_resolved_value(wb, sheet, cell):
    visited = set()
    return _resolve(wb, sheet, cell, visited)

def _resolve(wb, sheet, cell, visited):
    key = (sheet.title, cell.row, cell.column)
    if key in visited:
        return None
    visited.add(key)
    
    val = cell.value
    if val is None:
        return None
    val_str = str(val).strip()
    if val_str.startswith('='):
        target = val_str[1:].replace('$', '').strip()
        if '!' in target:
            parts = target.split('!', 1)
            target_sheet_name = parts[0].strip('\'"')
            target_cell_coord = parts[1]
            if target_sheet_name in wb.sheetnames:
                target_sheet = wb[target_sheet_name]
                try:
                    target_cell = target_sheet[target_cell_coord]
                    return _resolve(wb, target_sheet, target_cell, visited)
                except Exception:
                    pass
        else:
            try:
                target_cell = sheet[target]
                return _resolve(wb, sheet, target_cell, visited)
            except Exception:
                pass
    return val

def find_pqp_revision(wb):
    for sheetname in ['TitlePg', 'PQS', 'Index']:
        if sheetname in wb.sheetnames:
            sheet = wb[sheetname]
            for r in range(1, min(20, sheet.max_row + 1)):
                for c in range(1, min(20, sheet.max_column + 1)):
                    val = str(sheet.cell(row=r, column=c).value or '').upper()
                    if 'QF-SH-020' in val or 'QF-HH-020' in val:
                        rev_match = re.search(r'\bA\d+\b', val)
                        if rev_match:
                            return rev_match.group(0)
                        for offset in range(1, 4):
                            neighbor_val = str(sheet.cell(row=r, column=c+offset).value or '').upper()
                            rev_match = re.search(r'\bA\d+\b', neighbor_val)
                            if rev_match:
                                return rev_match.group(0)
    return "A8"

def find_dimensional_columns(sheet):
    cols = {
        "balloon": 2,
        "nominal": 4,
        "tol_up": 5,
        "tol_lo": 6,
        "unit": 14,
        "sample": 1,
        "meas": 9,
        "disp": 12
    }
    
    for r in range(20, 26):
        for c in range(1, min(30, sheet.max_column + 1)):
            val = str(sheet.cell(row=r, column=c).value or '').strip().lower()
            if 'balloon' in val or '尺寸#' in val:
                cols["balloon"] = c
            elif 'nominal' in val or '標準值' in val:
                cols["nominal"] = c
            elif 'upper tol' in val or 'upper\ntol' in val or '上差' in val:
                cols["tol_up"] = c
            elif 'lower tol' in val or 'lower\ntol' in val or '下差' in val:
                cols["tol_lo"] = c
            elif 'unit' in val or '單位' in val:
                cols["unit"] = c
            elif 'sample' in val or '樣品#' in val:
                cols["sample"] = c
            elif 'measurement method' in val or 'method' in val or '量測工具' in val or '量檢具' in val:
                cols["tool"] = c
            elif val == 'measurement' or val.startswith('measurement\n') or '量測值' in val or (('measurement' in val or 'measured' in val) and 'method' not in val):
                cols["meas"] = c
            elif val == 'disposition' or ('disposition' in val and '\n' not in val and '>' not in val) or '判定' in val:
                cols["disp"] = c
                
    return cols


def extract_pqs_vml_checkboxes(pqp_path=None, z=None):
    """
    Parses VML checkboxes and ctrlProps for the PQS sheet in an Excel workbook (.xlsx/.xlsm).
    Returns a dict mapping row (1-based) to dict:
    {
        'approved': bool,
        'conditional': bool,
        'rejected': bool,
        'na': bool
    }
    """
    should_close = False
    if z is None:
        if not pqp_path or not os.path.exists(pqp_path) or not zipfile.is_zipfile(pqp_path):
            return {}
        try:
            z = zipfile.ZipFile(pqp_path, 'r')
            should_close = True
        except Exception:
            return {}

    try:
        pqs_sheet_target = None
        if 'xl/workbook.xml' in z.namelist() and 'xl/_rels/workbook.xml.rels' in z.namelist():
            try:
                wb_xml = ET.fromstring(z.read('xl/workbook.xml'))
                wb_rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
                rid_to_target = {r.attrib['Id']: r.attrib['Target'] for r in wb_rels}
                for s in wb_xml.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
                    if s.attrib.get('name', '').strip().upper() == 'PQS':
                        rid = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                        if rid in rid_to_target:
                            target = rid_to_target[rid]
                            pqs_sheet_target = 'xl/' + target if not target.startswith('xl/') else target
                        break
            except Exception:
                pass
        
        if not pqs_sheet_target:
            pqs_sheet_target = 'xl/worksheets/sheet1.xml'
            
        pqs_rels_path = 'xl/worksheets/_rels/' + pqs_sheet_target.split('/')[-1] + '.rels'
        if pqs_rels_path not in z.namelist():
            return {}
            
        pqs_rels = ET.fromstring(z.read(pqs_rels_path))
        rid_to_target = {r.attrib['Id']: r.attrib['Target'] for r in pqs_rels}
        
        controls = {}
        if pqs_sheet_target in z.namelist():
            sheet_root = ET.fromstring(z.read(pqs_sheet_target))
            ns_main = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for c in sheet_root.findall('.//main:control', ns_main):
                shape_id = c.attrib.get('shapeId')
                rid = c.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                target = rid_to_target.get(rid, '')
                controls[shape_id] = {'target': target}

        for shape_id, info in controls.items():
            target = info.get('target', '')
            if target:
                prop_path = 'xl/' + target if not target.startswith('xl/') else target
                prop_path = prop_path.replace('worksheets/../', '')
                if prop_path in z.namelist():
                    try:
                        xml_data = z.read(prop_path).decode('utf-8', errors='ignore')
                        info['checked'] = ('checked="Checked"' in xml_data) or ('checked="1"' in xml_data)
                    except Exception:
                        info['checked'] = False
                else:
                    info['checked'] = False
            else:
                info['checked'] = False

        vml_targets = []
        for r in pqs_rels:
            if r.attrib.get('Type', '').endswith('/vmlDrawing'):
                t = r.attrib.get('Target', '')
                vml_path = 'xl/' + t if not t.startswith('xl/') else t
                vml_path = vml_path.replace('worksheets/../', '')
                if vml_path in z.namelist():
                    vml_targets.append(vml_path)
                    
        if not vml_targets and 'xl/drawings/vmlDrawing1.vml' in z.namelist():
            vml_targets.append('xl/drawings/vmlDrawing1.vml')

        rows_cbs = {}
        for vml_path in vml_targets:
            try:
                vml_text = z.read(vml_path).decode('utf-8', errors='ignore')
                shapes = re.findall(r'(<v:shape[^>]*>.*?</v:shape>)', vml_text, re.DOTALL)
                for s in shapes:
                    id_m = re.search(r'id="([^"]+)"', s)
                    shape_id = id_m.group(1) if id_m else None
                    anchor_m = re.search(r'<x:Anchor>([^<]+)</x:Anchor>', s)
                    anchor = anchor_m.group(1).strip() if anchor_m else ''
                    checked_m = re.search(r'<x:Checked>([^<]+)</x:Checked>', s)
                    
                    ctrl_info = controls.get(shape_id, {})
                    is_checked = ctrl_info.get('checked', False) or (checked_m is not None and checked_m.group(1) == '1')
                    
                    if anchor:
                        parts = [p.strip() for p in anchor.split(',')]
                        if len(parts) >= 8:
                            c_val = int(parts[0]) + 1
                            r_val = int(parts[2]) + 1
                            if r_val not in rows_cbs:
                                rows_cbs[r_val] = {}
                            rows_cbs[r_val][c_val] = {
                                'shape_id': shape_id,
                                'checked': is_checked
                            }
            except Exception:
                pass

        result = {}
        for r in range(50, 76):
            cbs = rows_cbs.get(r, {})
            app = any(c in [3, 4] and info['checked'] for c, info in cbs.items())
            cond = any(c in [7, 8] and info['checked'] for c, info in cbs.items())
            rej = any(c in [11, 12] and info['checked'] for c, info in cbs.items())
            na = any(c in [15, 16] and info['checked'] for c, info in cbs.items())
            result[r] = {
                'approved': app,
                'conditional': cond,
                'rejected': rej,
                'na': na
            }
        return result
    except Exception:
        return {}
    finally:
        if should_close:
            try:
                z.close()
            except Exception:
                pass


def run_checklist_alignment_check(wb, pqp_path=None):
    """
    Compares worksheet data presence against the markings in Index and PQS.
    Also audits VML Form Control Checkboxes in PQS, detects whether Approved and N/A
    are simultaneously checked for the same item, and checks 3-way alignment.
    """
    mismatches = []
    
    sheets_with_embedded = set()
    sheets_with_vector_drawings = set()
    vml_pqs_cbs = {}
    if pqp_path and os.path.exists(pqp_path):
        try:
            b_data = read_file_bytes_safe(pqp_path)
            z = zipfile.ZipFile(io.BytesIO(b_data))
            obj_to_sheets = get_embedded_to_sheets_map(z)
            for sheets in obj_to_sheets.values():
                for s in sheets:
                    sheets_with_embedded.add(s)

            # Map sheetId/rId to sheet name and find vector drawing XMLs (e.g. Process Flow Chart)
            wb_xml = ET.fromstring(z.read('xl/workbook.xml'))
            sheets_map = {}
            for s in wb_xml.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
                sheets_map[s.attrib['sheetId']] = (s.attrib['name'], s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'])
            wb_rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
            rId_to_target = {r.attrib['Id']: r.attrib['Target'] for r in wb_rels.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')}
            for s_id, (s_name, rId) in sheets_map.items():
                s_target = 'xl/' + rId_to_target.get(rId, '')
                parts = s_target.split('/')
                rels_path = '/'.join(parts[:-1]) + '/_rels/' + parts[-1] + '.rels'
                if rels_path in z.namelist():
                    s_rels_xml = ET.fromstring(z.read(rels_path))
                    for r in s_rels_xml.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                        if r.attrib['Type'].endswith('/drawing'):
                            d_target = r.attrib['Target']
                            d_path = 'xl/drawings/' + os.path.basename(d_target)
                            if d_path in z.namelist() and z.getinfo(d_path).file_size > 10000:
                                sheets_with_vector_drawings.add(s_name)
                                
            # Extract VML Form Control Checkboxes for PQS
            vml_pqs_cbs = extract_pqs_vml_checkboxes(z=z)
        except Exception:
            pass

    # 1. Check data presence for all worksheets
    sheet_status = {}
    for name in wb.sheetnames:
        sheet = wb[name]
        if name in ["PQS", "Index", "Revision History ", "Revision History"]:
            sheet_status[name] = True
            continue
        if name == "Drawing":
            sheet_status[name] = len(getattr(sheet, '_images', [])) > 1 or "Drawing" in sheets_with_embedded
            continue
            
        has_data = False
        
        if name.strip() == "Process Capability":
            # Check numeric subgroup data cells in rows 35..60
            for r in range(35, 60):
                for c in range(2, 16):
                    val = sheet.cell(row=r, column=c).value
                    if val is not None:
                        try:
                            float(str(val).strip())
                            has_data = True
                            break
                        except ValueError:
                            pass
                if has_data:
                    break

        elif name.strip() == "Packaging Test Results":
            # Test data is in rows 22, 24, 26, 28, 30 for cols 2..50
            has_data = False
            for r in [22, 24, 26, 28, 30]:
                for c in range(2, min(50, sheet.max_column + 1)):
                    val = sheet.cell(row=r, column=c).value
                    if val is not None:
                        val_str = str(val).strip()
                        if val_str != "" and not val_str.startswith("#"):
                            has_data = True
                            break
                if has_data:
                    break
                    
        elif name.strip() in ["Graphics First Article", "Thermoform First Article"]:
            for r in range(18, 101):
                val = sheet.cell(row=r, column=31).value
                if val is not None:
                    val_str = str(val).strip()
                    if val_str != "" and not val_str.startswith("#") and val_str.lower() not in ["acceptable", "accept"]:
                        has_data = True
                        break
                        
        elif name.strip().startswith("Variable R&R"):
            has_data = False
            for r in range(24, 33):
                for c in range(3, 13):
                    val = sheet.cell(row=r, column=c).value
                    if val is not None:
                        try:
                            float(str(val).replace("%", "").strip())
                            has_data = True
                            break
                        except ValueError:
                            pass
                if has_data:
                    break
        elif name.strip().startswith("Attribute R&R"):
            has_data = False
            for r in range(14, 44):
                for c in [4, 5, 7, 8, 10, 11]:
                    val = sheet.cell(row=r, column=c).value
                    if val is not None:
                        has_data = True
                        break
                if has_data:
                    break
        else:
            for r in range(16, min(100, sheet.max_row + 1)):
                for c in range(1, min(30, sheet.max_column + 1)):
                    val = sheet.cell(row=r, column=c).value
                    if val is not None:
                        val_str = str(val).strip()
                        if val_str != "" and not val_str.startswith("#") and not val_str.startswith("~"):
                            if "insert" not in val_str.lower() and "original test" not in val_str.lower() and "not applicable" not in val_str.lower():
                                has_data = True
                                break
                if has_data:
                    break
            if not has_data and getattr(sheet, '_images', []):
                for img in sheet._images:
                    anc = getattr(img, 'anchor', None)
                    anc_from = getattr(anc, '_from', None) if anc else None
                    row = getattr(anc_from, 'row', None) if anc_from else getattr(anc, 'row', None)
                    if row is None or row >= 10:
                        has_data = True
                        break
            if not has_data and (name in sheets_with_embedded or name.strip() in sheets_with_embedded or name in sheets_with_vector_drawings or name.strip() in sheets_with_vector_drawings):
                has_data = True
                    
        sheet_status[name] = has_data

    # 2. Get Index Worksheet Status
    index_status = {}
    index_row_map = {}
    if "Index" in wb.sheetnames:
        sheet = wb["Index"]
        for r in range(12, 32):
            tab_name = sheet.cell(row=r, column=2).value
            status = sheet.cell(row=r, column=12).value
            if tab_name:
                tab_str = str(tab_name).strip()
                index_status[tab_str] = str(status).strip() if status else None
                index_row_map[tab_str] = r

    # Check PQS header row 53 false checkbox check
    if vml_pqs_cbs.get(53, {}).get("approved") or vml_pqs_cbs.get(53, {}).get("na"):
        mismatches.append({
            "sheet": "PQS",
            "detail": "PQS 審核表頭標題列 (Row 53) 核取方塊被勾選（可能為操作誤觸），該列為標題，建議取消勾選。",
            "row": 53,
            "col": 3
        })

    # 3. Get PQS Worksheet Status (combining Cell Text and VML Form Control Checkboxes)
    pqs_status = {}
    if "PQS" in wb.sheetnames:
        sheet = wb["PQS"]
        for r in range(54, 73):
            req = sheet.cell(row=r, column=19).value
            cell_app = sheet.cell(row=r, column=3).value
            cell_cond = sheet.cell(row=r, column=7).value
            cell_rej = sheet.cell(row=r, column=11).value
            cell_na = sheet.cell(row=r, column=15).value
            
            vml_info = vml_pqs_cbs.get(r, {})
            
            def is_marked(cell_v, vml_v):
                if vml_v:
                    return True
                if cell_v is not None:
                    s = str(cell_v).strip().upper()
                    return s in ["X", "V", "1", "TRUE", "CHECKED", "✓"]
                return False
                
            app_marked = is_marked(cell_app, vml_info.get("approved", False))
            cond_marked = is_marked(cell_cond, vml_info.get("conditional", False))
            rej_marked = is_marked(cell_rej, vml_info.get("rejected", False))
            na_marked = is_marked(cell_na, vml_info.get("na", False))
            
            if req:
                pqs_status[r] = {
                    "text": str(req).strip(),
                    "approved": app_marked,
                    "conditional": cond_marked,
                    "rejected": rej_marked,
                    "na": na_marked
                }

    index_to_sheet = {
        "PQS": "PQS",
        "Drawing": "Drawing",
        "BOM-Matrices": "BOM - Matrices",
        "Material Certificates": "Material Certs",
        "DFM": "DFM",
        "FFFs": " FFF",
        "100% Dimensional Results": "100% Dimensional Results",
        "Process Capability": "Process Capability",
        "Process Flow Chart": "Process Flow Chart",
        "Process FMEA": "Process FMEA",
        "Control Plan": "Control Plan",
        "Inspection Plan": "Inspection Plan",
        "Variable R&R": "Variable R&R",
        "Attribute R&R": "Attribute R&R",
        "Finish Test Report": "Finish Test Reports",
        "Mechanical Test Reports": "Mechanical Test Reports",
        "Packaging Test Results": "Packaging Test Results",
        "Packaging Test Reports": "Packaging Test Reports",
        "Graphics First Article": "Graphics First Article",
        "Thermoform First Article": "Thermoform First Article "
    }

    pqs_row_to_sheet = {
        55: "Drawing",
        56: "BOM - Matrices",
        57: "Material Certs",
        58: "DFM",
        59: " FFF",
        60: "100% Dimensional Results",
        61: "Process Capability",
        62: "Process Flow Chart",
        63: "Process FMEA",
        64: "Control Plan",
        65: "Variable R&R",
        66: "Finish Test Reports",
        67: "Graphics First Article",
        68: "Thermoform First Article ",
        69: "Packaging Test Results",
        70: "Mechanical Test Reports",
    }

    def resolve_matching_sheets(target_name):
        t_clean = target_name.strip().lower()
        if t_clean == "100% dimensional results":
            return [s for s in wb.sheetnames if s.strip().lower().startswith("100% dimensional")]
        if t_clean == "variable r&r":
            return [s for s in wb.sheetnames if s.strip().lower().startswith("variable r&r")]
        if t_clean == "attribute r&r":
            return [s for s in wb.sheetnames if s.strip().lower().startswith("attribute r&r")]
        return [
            s for s in wb.sheetnames
            if s.strip().lower() == t_clean or s.strip().lower().rstrip('s') == t_clean.rstrip('s')
        ]

    # Verify Index
    for index_item, status in index_status.items():
        sheet_name = index_to_sheet.get(index_item, index_item)
        matched_sheets = resolve_matching_sheets(sheet_name)
        sheet_exists = len(matched_sheets) > 0
        actual_active = any(sheet_status.get(s, False) for s in matched_sheets)
                
        if not sheet_exists and index_item not in ["PQS", "Index"]:
            if status == "x":
                mismatches.append({
                    "sheet": "Index",
                    "detail": f"Index 標記為啟用 ('x')，但查無工作表 '{sheet_name}'。",
                    "row": index_row_map.get(index_item),
                    "col": 12
                })
            continue

        if status == "x" and not actual_active and index_item not in ["PQS", "Index"]:
            mismatches.append({
                "sheet": "Index",
                "detail": f"工作表 '{sheet_name}' 標記為啟用 ('x')，但內容為空或無有效數據。",
                "row": index_row_map.get(index_item),
                "col": 12
            })
        elif status in ["N/A", "NA"] and actual_active:
            mismatches.append({
                "sheet": "Index",
                "detail": f"工作表 '{sheet_name}' 標記為 'N/A'，但內部包含有效數據或圖檔。",
                "row": index_row_map.get(index_item),
                "col": 12
            })

    # 3.1 Check for simultaneous / conflicting marks on the same item in PQS
    total_pqs_marks = 0
    for r, info in pqs_status.items():
        req_text = info["text"]
        app_m = info["approved"]
        cond_m = info["conditional"]
        rej_m = info["rejected"]
        na_m = info["na"]
        
        if app_m or cond_m or rej_m or na_m:
            total_pqs_marks += 1
            
        # Strict check: Are BOTH Approved and N/A checked for the same item?
        if app_m and na_m:
            mismatches.append({
                "sheet": "PQS",
                "detail": f"❌ **FAIL (衝突勾選)**: 第 {r} 列 '{req_text}' 同時勾選了 'Approved' 與 'N/A'，狀態嚴重互斥！",
                "row": r,
                "col": 3
            })
            
        # Check for multiple statuses checked (e.g. Approved + Rejected)
        st_list = []
        if app_m: st_list.append("Approved")
        if cond_m: st_list.append("Conditional")
        if rej_m: st_list.append("Rejected")
        if na_m: st_list.append("N/A")
        if len(st_list) > 1 and not (app_m and na_m):
            mismatches.append({
                "sheet": "PQS",
                "detail": f"⚠️ **WARNING (多重勾選)**: 第 {r} 列 '{req_text}' 同時勾選了多種狀態 ({'/'.join(st_list)})！",
                "row": r,
                "col": 3
            })

    if total_pqs_marks == 0 and pqs_status:
        mismatches.append({
            "sheet": "PQS",
            "detail": "PQS 審核檢核表 (Rows 54-72) 全數欄位尚未勾選 (Approved / N/A 均為空白)，目前為供應商初版送審狀態，待 ASSA ABLOY 品保工程師審核簽核。",
            "row": 54,
            "col": 3
        })
    else:
        # 3.2 3-Way Consistency Check (PQS vs Index vs Worksheet data)
        for r, info in pqs_status.items():
            sheet_name = pqs_row_to_sheet.get(r)
            if not sheet_name:
                continue
                
            if r == 55:  # Drawing
                actual_active = sheet_status.get("Drawing", False) or "Drawing" in sheets_with_embedded or len(getattr(wb.get("Drawing", None), '_images', [])) > 0 or (index_status.get("Drawing") == "x")
            elif r == 56:  # BOM / Matrices
                actual_active = (index_status.get("BOM-Matrices") == "x") and any(sheet_status.get(s, False) for s in resolve_matching_sheets("BOM - Matrices"))
            elif r == 65:  # Gauge R&R
                actual_active = any(sheet_status.get(s, False) for s in wb.sheetnames if "R&R" in s or "Variable R&R" in s or "Attribute R&R" in s) or (index_status.get("Variable R&R") == "x") or (index_status.get("Attribute R&R") == "x")
            elif r == 69:  # Packaging
                actual_active = any(sheet_status.get(s, False) for s in resolve_matching_sheets("Packaging Test Results") + resolve_matching_sheets("Packaging Test Reports"))
            else:
                actual_active = any(sheet_status.get(s, False) for s in resolve_matching_sheets(sheet_name))
                
            approved_marked = info["approved"]
            na_marked = info["na"]
            
            if actual_active:
                if not approved_marked and not info["conditional"]:
                    mismatches.append({
                        "sheet": "PQS",
                        "detail": f"第 {r} 列 '{info['text']}' 在活頁簿中有資料 (或 Index 標記為啟用)，但 PQS 未勾選 Approved。",
                        "row": r,
                        "col": 3
                    })
                if na_marked:
                    mismatches.append({
                        "sheet": "PQS",
                        "detail": f"❌ **FAIL (誤勾 N/A)**: 第 {r} 列 '{info['text']}' 在活頁簿中有資料 (且 Index 標記為啟用)，但 PQS 錯誤勾選為 N/A！",
                        "row": r,
                        "col": 15
                    })
            else:
                if approved_marked:
                    mismatches.append({
                        "sheet": "PQS",
                        "detail": f"❌ **FAIL (誤勾 Approved)**: 第 {r} 列 '{info['text']}' 在活頁簿中無有效資料 (Index 標記為 N/A)，但 PQS 錯誤勾選為 Approved！",
                        "row": r,
                        "col": 3
                    })
                if not na_marked and (r != 72):  # Row 72 Pilot Run can remain blank
                    mismatches.append({
                        "sheet": "PQS",
                        "detail": f"第 {r} 列 '{info['text']}' 無有效資料 (Index 標記為 N/A)，但 PQS 未勾選 N/A。",
                        "row": r,
                        "col": 15
                    })
                
    return mismatches

def run_metadata_audit(wb, wb_formulas, expected_dwg, expected_rev, expected_desc=None, expected_pn=None):
    mismatches = []
    
    mapping = {
        "Part Number": ["part number", "part #", "part no", "part no.", "item number", "item no"],
        "Part Description": ["part description", "description", "part desc"],
        "Drawing Number": ["drawing number", "drawing #", "drawing no", "drawing no.", "dwg no", "dwg #", "dwg no."],
        "Revision": ["revision", "rev", "rev.", "drawing rev", "dwg rev"]
    }
    
    def get_resolved_val(sheet_name, r, c):
        if wb_formulas and sheet_name in wb_formulas.sheetnames:
            cell_formula = wb_formulas[sheet_name].cell(row=r, column=c)
            resolved = get_resolved_value(wb_formulas, wb_formulas[sheet_name], cell_formula)
            if resolved is not None:
                resolved_str = str(resolved).strip()
                if not resolved_str.startswith('='):
                    return resolved
        if sheet_name in wb.sheetnames:
            return wb[sheet_name].cell(row=r, column=c).value
    expected_pn_found = None
    expected_desc_found = None
    header_sheet_name = 'TitlePg' if 'TitlePg' in wb.sheetnames else ('PQS' if 'PQS' in wb.sheetnames else None)
    if header_sheet_name:
        h_sheet = wb[header_sheet_name]
        for r in range(1, min(25, h_sheet.max_row + 1)):
            for c in range(1, min(25, h_sheet.max_column + 1)):
                val_str = str(h_sheet.cell(row=r, column=c).value or '').strip()
                val_lower = val_str.lower()
                for key, search_terms in mapping.items():
                    if any(term == val_lower or val_lower.startswith(term + ":") or val_lower.startswith(term + " ") for term in search_terms):
                        assoc_val = None
                        if ":" in val_str:
                            assoc_val = val_str.split(":", 1)[1].strip()
                        if not assoc_val:
                            for col_offset in range(1, 10):
                                if c + col_offset <= h_sheet.max_column:
                                    peek = h_sheet.cell(row=r, column=c+col_offset).value
                                    if peek is not None and str(peek).strip():
                                        peek_str = str(peek).strip()
                                        if not any(any(t == peek_str.lower() or peek_str.lower().startswith(t + ":") for t in st) for st in mapping.values()):
                                            assoc_val = peek_str
                                            break
                        if assoc_val:
                            if key == "Part Number":
                                expected_pn_found = assoc_val
                                if expected_pn and robust_norm(str(assoc_val)) != robust_norm(str(expected_pn)):
                                    mismatches.append({
                                        "sheet": header_sheet_name,
                                        "field": "Part Number (PQS 封面)",
                                        "expected": expected_pn,
                                        "found": assoc_val,
                                        "reason": f"PQS 封面 Part Number ('{assoc_val}') 與工程圖面料號 ('{expected_pn}') 不符！"
                                    })
                                elif not expected_pn:
                                    expected_pn = assoc_val
                            elif key == "Part Description":
                                expected_desc_found = assoc_val
                                if expected_desc and str(expected_desc).strip().upper() not in str(assoc_val).strip().upper() and str(assoc_val).strip().upper() not in str(expected_desc).strip().upper():
                                    mismatches.append({
                                        "sheet": header_sheet_name,
                                        "field": "Part Description (PQS 封面)",
                                        "expected": expected_desc,
                                        "found": assoc_val,
                                        "reason": f"PQS 封面 Part Description ('{assoc_val}') 與工程圖面品名 ('{expected_desc}') 不符！"
                                    })
                                elif not expected_desc:
                                    expected_desc = assoc_val
                            elif key == "Drawing Number":
                                if expected_dwg and robust_norm(str(assoc_val)) != robust_norm(str(expected_dwg)):
                                    mismatches.append({
                                        "sheet": header_sheet_name,
                                        "field": "Drawing Number (PQS 封面)",
                                        "expected": expected_dwg,
                                        "found": assoc_val,
                                        "reason": f"PQS 封面 Drawing Number ('{assoc_val}') 與工程圖面圖號 ('{expected_dwg}') 不符！"
                                    })
                                elif not expected_dwg:
                                    expected_dwg = assoc_val
                            elif key == "Revision":
                                if expected_rev and str(assoc_val).strip() != str(expected_rev).strip():
                                    mismatches.append({
                                        "sheet": header_sheet_name,
                                        "field": "Revision (PQS 封面)",
                                        "expected": expected_rev,
                                        "found": assoc_val,
                                        "reason": f"PQS 封面 Revision ('{assoc_val}') 與工程圖面版次 ('{expected_rev}') 不符！"
                                    })
                                elif not expected_rev:
                                    expected_rev = assoc_val
        
    active_sheets = [s for s in wb.worksheets if s.sheet_state == 'visible']
    
    for sheet in active_sheets:
        sheet_name = sheet.title
        name_lower = sheet_name.strip().lower()
        if "index" in name_lower or "revision history" in name_lower or name_lower.endswith("pqs"):
            continue
            
        max_r = min(40, sheet.max_row)
        max_c = min(40, sheet.max_column)
        sheet_meta = {}
        
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                cell_val = sheet.cell(row=r, column=c).value
                if cell_val is not None:
                    val_str = str(cell_val).strip()
                    val_lower = val_str.lower()
                    matched_key = None
                    for key, search_terms in mapping.items():
                        if any(term == val_lower or val_lower.startswith(term + ":") or val_lower.startswith(term + " ") for term in search_terms):
                            matched_key = key
                            break
                    if matched_key:
                        assoc_val = None
                        if ":" in val_str:
                            assoc_val = val_str.split(":", 1)[1].strip()
                        if not assoc_val:
                            for col_offset in range(1, 15):
                                if c + col_offset <= sheet.max_column:
                                    f_sheet = wb_formulas[sheet_name] if wb_formulas and sheet_name in wb_formulas.sheetnames else sheet
                                    peek_val = f_sheet.cell(row=r, column=c+col_offset).value
                                    if peek_val is not None:
                                        peek_val_str = str(peek_val).strip()
                                        is_other_key = False
                                        for k, s_terms in mapping.items():
                                            if any(term == peek_val_str.lower() or peek_val_str.lower().startswith(term + ":") or peek_val_str.lower().startswith(term + " ") for term in s_terms):
                                                is_other_key = True
                                                break
                                        if is_other_key:
                                            break
                                            
                                        resolved_peek = get_resolved_val(sheet_name, r, c+col_offset)
                                        if resolved_peek is not None:
                                            assoc_val = str(resolved_peek).strip()
                                            break
                        if assoc_val:
                            if matched_key == "Revision" and ("effective" in assoc_val.lower() or "date" in assoc_val.lower()):
                                continue
                            if matched_key not in sheet_meta:
                                sheet_meta[matched_key] = assoc_val
                            
        expectations = {
            "Part Number": expected_pn,
            "Part Description": expected_desc,
            "Drawing Number": expected_dwg,
            "Revision": expected_rev
        }
        
        for key, exp_val in expectations.items():
            if exp_val is not None:
                exp_str = str(exp_val).strip().lower()
                found_val = sheet_meta.get(key)
                if found_val:
                    found_str = str(found_val).strip().lower()
                    if found_str != exp_str:
                        try:
                            if float(found_str) == float(exp_str):
                                continue
                        except ValueError:
                            pass
                        if found_str.strip() == exp_str.strip():
                            continue
                            
                        # Support multi-part / family parts (split by newline or semicolon)
                        exp_parts = [p.strip() for p in exp_str.splitlines() if p.strip()]
                        if found_str.strip() in exp_parts:
                            continue
                        if any(found_str.strip() == p for p in exp_parts):
                            continue
                        # Also test if found_str is in comma-separated PN list
                        exp_comma_parts = [p.strip() for p in exp_str.replace('\n', ',').split(',') if p.strip()]
                        if found_str.strip() in exp_comma_parts:
                            continue
                            
                        ext_note = ""
                        if wb_formulas and sheet_name in wb_formulas.sheetnames:
                            f_ws = wb_formulas[sheet_name]
                            for r_chk in range(1, 15):
                                for c_chk in range(1, 15):
                                    f_val = str(f_ws.cell(row=r_chk, column=c_chk).value or '')
                                    if f_val.startswith('='):
                                        m_ext = re.search(r'\[(\d+)\]([^!]+)!', f_val)
                                        if m_ext:
                                            ext_note = f" (⚠️ 公式 '{f_val}' 含外部活頁簿索引 '[{m_ext.group(1)}]'，導致顯示來源舊值；請移除 '[{m_ext.group(1)}]')"
                                            break
                                if ext_note:
                                    break
                                    
                        mismatches.append({
                            "sheet": sheet_name,
                            "field": key + ext_note,
                            "expected": exp_val,
                            "found": found_val
                        })
                else:
                    if name_lower in ["control plan", "inspection plan", "process fmea", "variable r&r", "gr&r"]:
                        mismatches.append({
                            "sheet": sheet_name,
                            "field": key,
                            "expected": exp_val,
                            "found": "[未找到欄位/公式連結空白]"
                        })

    # Audit hidden sheets for residual foreign project data
    hidden_sheets = [s for s in wb.worksheets if s.sheet_state != 'visible']
    for sheet in hidden_sheets:
        h_name = sheet.title
        max_r = min(25, sheet.max_row)
        max_c = min(25, sheet.max_column)
        foreign_tokens = []
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                cv = sheet.cell(row=r, column=c).value
                if cv:
                    cv_s = str(cv).strip()
                    m_pn = re.search(r'\b(KPZ[A-Z0-9]+|\d{3}-\d{5}-\d{3}|\d{5}-\d{3})\b', cv_s)
                    if m_pn and (not expected_dwg or robust_norm(m_pn.group(1)) != robust_norm(str(expected_dwg))):
                        foreign_tokens.append(m_pn.group(1))
        if foreign_tokens:
            u_pns = list(dict.fromkeys(foreign_tokens))[:3]
            mismatches.append({
                "sheet": f"{h_name} [隱藏工作表]",
                "field": "Hidden Foreign Sheet",
                "expected": f"無多餘舊專案隱藏分頁 ({expected_dwg})",
                "found": f"殘留料號: {', '.join(u_pns)}",
                "reason": f"工作表 '{h_name}' 被隱藏但包含前專案料號 ({', '.join(u_pns)})，造成活頁簿膨脹至 88MB，應徹底刪除以防混淆與洩密。"
            })

    return mismatches

def run_control_plan_audit(wb):
    cp_sheets = [s for s in wb.sheetnames if "control plan" in s.lower()]
    if not cp_sheets:
        return []
    sheet = wb[cp_sheets[0]]
    mismatches = []
    
    for r in range(16, sheet.max_row + 1):
        proc_desc = sheet.cell(row=r, column=3).value
        if proc_desc:
            desc_str = str(proc_desc).lower()
            
            if "adaptor" in desc_str and ("screw" in desc_str or "fix" in desc_str) and "back plate" not in desc_str:
                torque_match = re.search(r'torque\s+is\s+([\d\.]+)', desc_str)
                if torque_match:
                    val = float(torque_match.group(1))
                    if val != 1.4:
                        mismatches.append({
                            "row": r,
                            "step": sheet.cell(row=r, column=2).value,
                            "description": proc_desc,
                            "found_torque": val,
                            "expected_torque": 1.4,
                            "reason": "Rose to adapter 螺絲扭力應為 1.4 ± 0.1 N-m (對應 5057209 Flag Note 5)"
                        })
                        
            if "back plate" in desc_str and ("screw" in desc_str or "fix" in desc_str):
                torque_match = re.search(r'torque\s+is\s+([\d\.]+)', desc_str)
                if torque_match:
                    val = float(torque_match.group(1))
                    if val != 1.0:
                        mismatches.append({
                            "row": r,
                            "step": sheet.cell(row=r, column=2).value,
                            "description": proc_desc,
                            "found_torque": val,
                            "expected_torque": 1.0,
                            "reason": "Back plate 螺絲扭力應為 1.0 ± 0.1 N-m (對應 5057209 Flag Note 4)"
                        })
                        
            if "outer mount plate" in desc_str and ("screw" in desc_str or "fix" in desc_str):
                torque_match = re.search(r'torque\s+is\s+([\d\.]+)', desc_str)
                if torque_match:
                    val = float(torque_match.group(1))
                    if val != 0.84:
                        mismatches.append({
                            "row": r,
                            "step": sheet.cell(row=r, column=2).value,
                            "description": proc_desc,
                            "found_torque": val,
                            "expected_torque": 0.84,
                            "reason": "Outer mount plate 螺絲扭力應為 0.84 ± 0.1 N-m (對應 5052389 Note 4)"
                        })
                        
    return mismatches

def run_control_plan_ctq_sc_audit(wb, dwg_specs=None):
    """
    Audits coverage and specificity of Critical To Quality (CTQ) and Significant Characteristics (SC)
    in the Control Plan worksheet, cross-referencing against:
    - 100% Dimensional Results (SC/CTQ column)
    - Variable R&R-1..N (SC/CTQ characteristics and gages)
    - Process Capability sheets
    - Drawing process flag notes (e.g. Paint thickness, Heat treatment)
    """
    results = {
        "status": "PASS",
        "sc_ctq_items": [],
        "sc_ctq_dict": {},
        "all_dim_balloons": {},
        "control_plan_entries": [],
        "findings": []
    }
    
    # 1. Extract SC/CTQ items from 100% Dimensional Results
    dim_sc_items = {}
    dim_sheets = [s for s in wb.sheetnames if "Dimensional" in s and wb[s].sheet_state == 'visible']
    for sname in dim_sheets:
        ws = wb[sname]
        cols = find_dimensional_columns(ws)
        # Find SC/CTQ column
        sc_col = 3
        for c in range(1, 15):
            val = str(ws.cell(row=24, column=c).value or '').upper()
            if 'SC' in val or 'CTQ' in val or 'SIGNIFICANT' in val:
                sc_col = c
                break
                
        for r in range(25, ws.max_row + 1):
            b_val = ws.cell(row=r, column=cols.get("balloon", 2)).value
            sc_val = ws.cell(row=r, column=sc_col).value
            nom = ws.cell(row=r, column=cols.get("nominal", 4)).value
            up = ws.cell(row=r, column=cols.get("tol_up", 5)).value
            lo = ws.cell(row=r, column=cols.get("tol_lo", 6)).value
            unit = ws.cell(row=r, column=cols.get("unit", 14)).value
            tool = ws.cell(row=r, column=13).value
            
            if sc_val and any(k in str(sc_val).upper() for k in ['SC', 'CTQ', 'SIGNIFICANT', 'CRITICAL']):
                if b_val is not None and b_val not in dim_sc_items:
                    dim_sc_items[b_val] = {
                        "balloon": b_val,
                        "type": str(sc_val).strip(),
                        "nominal": nom,
                        "tol_up": up,
                        "tol_lo": lo,
                        "unit": unit or "mm",
                        "tool": tool,
                        "sheet": sname
                    }
                    
    # 2. Extract Variable R&R sheets info
    var_rr_items = []
    for sname in wb.sheetnames:
        if 'Variable R&R' in sname and wb[sname].sheet_state == 'visible':
            ws = wb[sname]
            char_type = ws.cell(row=12, column=4).value or ws.cell(row=4, column=18).value
            spec_val = ws.cell(row=13, column=4).value or ws.cell(row=4, column=25).value
            tol_up_val = ws.cell(row=14, column=4).value
            tol_lo_val = ws.cell(row=15, column=4).value
            gage_name = ws.cell(row=12, column=8).value or ws.cell(row=4, column=8).value
            var_rr_items.append({
                "sheet": sname,
                "char_type": char_type,
                "spec": spec_val,
                "tol_up": tol_up_val,
                "tol_lo": tol_lo_val,
                "gage": gage_name
            })
            
    # 3. Extract Control Plan rows
    cp_sheets = [s for s in wb.sheetnames if "control plan" in s.lower() and wb[s].sheet_state == 'visible']
    if not cp_sheets:
        results["status"] = "WARNING"
        results["findings"].append("可見工作表中未找到 [Control Plan] 管制計畫工作表。")
        return results
        
    ws_cp = wb[cp_sheets[0]]
    
    # Dynamically find column positions
    col_proc_num = 2
    col_proc_name = 3
    col_class = 6
    col_prod_char = 7
    col_proc_char = 8
    col_spec = 9
    col_eval_tech = 10
    col_sample_size = 11
    col_freq = 12
    col_reaction = 15
    
    for r in range(12, 18):
        for c in range(1, ws_cp.max_column + 1):
            val = str(ws_cp.cell(row=r, column=c).value or '').upper()
            if 'PROCESS NUMBER' in val or 'PROCESS #' in val:
                col_proc_num = c
            elif 'PROCESS NAME' in val or 'OPERATION DESCRIPTION' in val:
                col_proc_name = c
            elif 'CLASSIFICATION' in val or 'TIB/CTQ/SC' in val or 'SPECIAL CHAR' in val:
                col_class = c
            elif 'PRODUCT' in val and 'CHARACTERISTIC' in str(ws_cp.cell(row=r-1, column=c).value or '').upper():
                col_prod_char = c
            elif 'PROCESS' in val and 'CHARACTERISTIC' in str(ws_cp.cell(row=r-1, column=c).value or '').upper():
                col_proc_char = c
            elif 'SPECIFICATION' in val or 'TOLERANCE' in val or 'SPEC' in val:
                col_spec = c
            elif 'EVALUATION' in val or 'MEASUREMENT TECHNIQUE' in val:
                col_eval_tech = c
            elif 'SAMPLE SIZE' in val:
                col_sample_size = c
            elif 'SAMPLE FREQ' in val or 'FREQUENCY' in val:
                col_freq = c
            elif 'REACTION' in val or 'REACTION PLAN' in val:
                col_reaction = c

    cp_rows = []
    for r in range(16, ws_cp.max_row + 1):
        p_num = ws_cp.cell(row=r, column=col_proc_num).value
        p_name = ws_cp.cell(row=r, column=col_proc_name).value
        special_char = ws_cp.cell(row=r, column=col_class).value
        prod_char = ws_cp.cell(row=r, column=col_prod_char).value
        proc_char = ws_cp.cell(row=r, column=col_proc_char).value
        spec = ws_cp.cell(row=r, column=col_spec).value
        eval_tech = ws_cp.cell(row=r, column=col_eval_tech).value
        sample_size = ws_cp.cell(row=r, column=col_sample_size).value
        freq = ws_cp.cell(row=r, column=col_freq).value
        reaction = ws_cp.cell(row=r, column=col_reaction).value
        
        if any([p_num, p_name, special_char, prod_char, spec, reaction]):
            cp_rows.append({
                "row": r,
                "proc_num": p_num,
                "proc_name": str(p_name or '').strip(),
                "special_char": str(special_char or '').strip(),
                "prod_char": str(prod_char or '').strip(),
                "proc_char": str(proc_char or '').strip(),
                "spec": str(spec or '').strip(),
                "eval_tech": str(eval_tech or '').strip(),
                "sample_size": str(sample_size or '').strip(),
                "freq": str(freq or '').strip(),
                "reaction_plan": str(reaction or '').strip()
            })
            
    results["sc_ctq_items"] = list(dim_sc_items.values())
    dwg_sc_notes = dwg_specs.get("sc_notes", []) if dwg_specs else []
    dwg_sc_features = dwg_specs.get("sc_features", []) if dwg_specs else []
    results["dwg_sc_notes"] = dwg_sc_notes
    results["dwg_sc_features"] = dwg_sc_features
    results["control_plan_entries"] = cp_rows

    # 4.0 Cross-check drawing SC diamond / circled letter features against 100% Dimensional Results
    target_sc_balloons = set()
    for b_val in dim_sc_items.keys():
        try:
            target_sc_balloons.add(int(float(str(b_val).strip())))
        except Exception:
            target_sc_balloons.add(b_val)

    all_dim_balloons = {}
    if dim_sheets:
        ws_dim = wb[dim_sheets[0]]
        cols_dim = find_dimensional_columns(ws_dim)
        sc_c_idx = 3
        for c in range(1, 15):
            val = str(ws_dim.cell(row=24, column=c).value or '').upper()
            if 'SC' in val or 'CTQ' in val or 'SIGNIFICANT' in val:
                sc_c_idx = c
                break

        for r in range(25, ws_dim.max_row + 1):
            b_val = ws_dim.cell(row=r, column=cols_dim.get("balloon", 2)).value
            sc_val = ws_dim.cell(row=r, column=sc_c_idx).value
            nom = ws_dim.cell(row=r, column=cols_dim.get("nominal", 4)).value
            up = ws_dim.cell(row=r, column=cols_dim.get("tol_up", 5)).value
            lo = ws_dim.cell(row=r, column=cols_dim.get("tol_lo", 6)).value
            if b_val is not None:
                try:
                    b_num = int(float(str(b_val).strip()))
                except Exception:
                    b_num = b_val
                if b_num not in all_dim_balloons:
                    all_dim_balloons[b_num] = {
                        "row": r,
                        "nominal": nom,
                        "tol_up": float(up) if isinstance(up, (int, float)) else 0.0,
                        "tol_lo": float(lo) if isinstance(lo, (int, float)) else 0.0,
                        "sc_val": str(sc_val or '').strip(),
                        "is_sc": bool(sc_val and any(k in str(sc_val).upper() for k in ['SC', 'CTQ', 'SIGNIFICANT', 'CRITICAL']))
                    }
                elif sc_val and any(k in str(sc_val).upper() for k in ['SC', 'CTQ', 'SIGNIFICANT', 'CRITICAL']):
                    all_dim_balloons[b_num]["is_sc"] = True

    if dwg_sc_features and dim_sheets:
        for feat in dwg_sc_features:
            f_nom = feat.get("nominal")
            f_letter = feat.get("letter")
            nearby_tol_str = feat.get("nearby_tol_str", "")
            if f_nom is None:
                continue
            
            candidates = []
            for b, info in all_dim_balloons.items():
                if info.get("nominal") is None:
                    continue
                try:
                    val = float(str(info["nominal"]).strip())
                    if abs(val - float(f_nom)) < 0.08:
                        candidates.append(b)
                except (ValueError, TypeError):
                    continue
            nearby_nums = []
            for tok in re.findall(r'[-+]?\d*\.?\d+', nearby_tol_str):
                try:
                    nearby_nums.append(abs(float(tok)))
                except ValueError:
                    pass

            best_b = None
            best_score = -999
            for b in candidates:
                score = 0
                info = all_dim_balloons[b]
                nom_diff = abs(float(info["nominal"]) - float(f_nom))
                if nom_diff < 1e-4:
                    score += 50
                elif nom_diff < 0.005:
                    score += 20
                else:
                    score -= int(nom_diff * 100)
                
                up_val = abs(info.get("tol_up", 0.0))
                lo_val = abs(info.get("tol_lo", 0.0))
                for val in [up_val, lo_val]:
                    if val != 0:
                        if any(abs(val - n) < 1e-4 for n in nearby_nums):
                            score += 5
                    elif val == 0:
                        if any(abs(n) < 1e-4 for n in nearby_nums) or '0' in nearby_tol_str.split():
                            score += 3
                if info.get("is_sc"):
                    score += 2
                if score > best_score:
                    best_score = score
                    best_b = b

            if best_b is not None:
                matched_b = best_b
                feat["matched_balloon"] = matched_b
                target_sc_balloons.add(matched_b)
                d_info = all_dim_balloons[matched_b]
                if not d_info["is_sc"]:
                    results["status"] = "FAIL"
                    results["findings"].append(
                        f"❌ **CRITICAL OMISSION (圖面 SC/CTQ 尺寸表漏標)**: 圖面特徵代號 `{f_letter}` (規格: {f_nom}) 在工程圖面標示為 {feat.get('type', 'SC')} 重點管制特徵，但在 [100% Dimensional Results] 中 Balloon #{matched_b} 之 SC/CTQ 欄位為空白，未標註！"
                    )
                elif feat.get("type") == "CTQ":
                    pqp_char_type = str(d_info.get("sc_val", "")).strip().upper()
                    if "CTQ" not in pqp_char_type:
                        results["status"] = "FAIL"
                        results["findings"].append(
                            f"❌ **CRITICAL MISCLASSIFICATION (CTQ 誤降為 SC)**: 圖面尺寸 Balloon #{matched_b} (規格: {f_nom}) 在工程圖面上標註為十字貫穿菱形 `⯐` (CTQ 關鍵品質特性，依 Note 7 強制要求 Cpk >= 1.33)，但在 [100% Dimensional Results] 中被標註為 `{d_info.get('sc_val')}` (誤降為一般重要特性 SC)！這將導致漏做必備之製程能力研究 (Process Capability Study)。"
                        )

    # 4.05 Check Variable R&R Completeness for all SC Balloons
    active_var_rr_sheets = [s for s in wb.sheetnames if 'variable r&r' in s.lower() and wb[s].sheet_state == 'visible']
    for sc_b in sorted(list(target_sc_balloons), key=lambda x: (isinstance(x, str), x)):
        has_rr_sheet = False
        for sname in active_var_rr_sheets:
            m_s = re.search(r'\((\d+)\)', sname)
            if m_s and str(sc_b) == m_s.group(1):
                has_rr_sheet = True
                break
            elif re.search(r'\b' + str(sc_b) + r'\b', sname):
                has_rr_sheet = True
                break
        if not has_rr_sheet:
            # Fallback: Check if any active Variable R&R sheet has matching specification in row 11~15
            sc_nom_val = None
            if sc_b in dim_sc_items and dim_sc_items[sc_b].get("nominal") is not None:
                try:
                    sc_nom_val = float(dim_sc_items[sc_b]["nominal"])
                except Exception:
                    pass
            elif dwg_sc_features:
                for f in dwg_sc_features:
                    if f.get("matched_balloon") == sc_b and f.get("nominal") is not None:
                        try:
                            sc_nom_val = float(f.get("nominal"))
                        except Exception:
                            pass
                        break
            if sc_nom_val is not None:
                for sname in active_var_rr_sheets:
                    ws_rr = wb[sname]
                    spec_val = None
                    for r_chk in range(11, 16):
                        c2_txt = str(ws_rr.cell(row=r_chk, column=2).value or "").strip().lower()
                        if "specification" in c2_txt or "spec" in c2_txt:
                            val = ws_rr.cell(row=r_chk, column=4).value
                            if val is not None:
                                try:
                                    spec_val = float(val)
                                    break
                                except Exception:
                                    pass
                    if spec_val is None:
                        val = ws_rr.cell(row=13, column=4).value
                        if val is not None:
                            try:
                                spec_val = float(val)
                            except Exception:
                                pass
                    if spec_val is not None and abs(spec_val - sc_nom_val) < 0.001:
                        has_rr_sheet = True
                        results["findings"].append(
                            f"ℹ️ **GAGE R&R MATCHED BY SPEC**: Balloon #{sc_b} (規格 {sc_nom_val}) 對應至 [{sname}]（依量具規格成功匹配，建議分頁名稱改為標註球標號如 Variable R&R ({sc_b}) 以提升辨識度）。"
                        )
                        break

        if not has_rr_sheet:
            results["status"] = "FAIL"
            nom_display = ""
            if sc_b in dim_sc_items:
                nom_display = f" (規格: {dim_sc_items[sc_b]['nominal']})"
            elif dwg_sc_features:
                for f in dwg_sc_features:
                    if f.get("matched_balloon") == sc_b:
                        nom_display = f" (規格: {f.get('nominal')})"
                        break
            is_visual_or_attr = False
            item_info = dim_sc_items.get(sc_b, {})
            nom_val_str = str(item_info.get("nominal", "")).strip()
            tool_val_str = str(item_info.get("tool", "")).strip().lower()
            if "visual" in tool_val_str or any(k in nom_val_str.lower() for k in ["print", "sticker", "material", "artwork", "appearance", "color"]):
                is_visual_or_attr = True

            if is_visual_or_attr:
                results["findings"].append(
                    f"❌ **CRITICAL GAGE R&R MISSING (漏做 Variable R&R)**: Balloon #{sc_b}{nom_display} 標註為 SC 重點特徵，但在活頁簿中完全未找到對應的 [Variable R&R ({sc_b})] 量具研究分頁！"
                    f"\n   *(💡 建議：此項為目視外觀/材質特性 (Visual)，工程圖面並無尺寸 SC 菱形。若非計量尺寸管制，建議在尺寸表取消 SC 標記，或改提供 Attribute MSA 計數型量具研究)*"
                )
            else:
                results["findings"].append(
                    f"❌ **CRITICAL GAGE R&R MISSING (漏做 Variable R&R)**: Balloon #{sc_b}{nom_display} 標註為 SC 重點特徵，但在活頁簿中完全未找到對應的 [Variable R&R ({sc_b})] 量具研究分頁！"
                )
    
    # 4. Check Traceability and Generic Descriptions for 100% Dimensional SCs
    sc_cp_rows = [row for row in cp_rows if any(k in row["special_char"].upper() for k in ['SC', 'CTQ', 'SIGNIFICANT', 'CRITICAL'])]
    
    if dim_sc_items and not sc_cp_rows:
        results["status"] = "FAIL"
        results["findings"].append(f"❌ **FAIL**: 圖面/尺寸表已定義 {len(dim_sc_items)} 項 SC/CTQ 特徵，但 Control Plan 中完全未標註任何 SC/CTQ 特殊特性工序。")
    elif dim_sc_items:
        # Check if dimensions are explicitly listed or generic
        generic_specs = []
        for row in sc_cp_rows:
            spec_txt = row["spec"].lower()
            if any(phrase in spec_txt for phrase in ["size meets", "size and appearance", "meets drawing", "meet specifications"]) and not any(char.isdigit() for char in row["spec"]):
                generic_specs.append(f"工序 {row['proc_num']} [{row['proc_name']}] (Row {row['row']})")
                
        if generic_specs:
            results["findings"].append(
                f"⚠️ **規格具體性提醒**：Control Plan 在 {', '.join(generic_specs)} 標註了 SC 特性，但規格欄填寫概括性字眼（如 '{sc_cp_rows[0]['spec']}'），建議具體列出各 Balloon 尺寸（如 " +
                ", ".join([f"Balloon #{d['balloon']}: {d['nominal']} (+{d['tol_up']}/-{d['tol_lo']} {d['unit']})" for d in dim_sc_items.values()]) + "）以符合嚴格審查標準。"
            )
            
        # Check Gage/Tool alignment
        cp_tools = " ".join([r["eval_tech"].lower() for r in sc_cp_rows])
        for var_item in var_rr_items:
            g_name = str(var_item.get("gage") or "").strip()
            if g_name and g_name.lower() not in cp_tools:
                results["findings"].append(
                    f"⚠️ **檢驗量具一致性提醒**：工作表 [{var_item['sheet']}] 指定量具為 `{g_name}` (規格: {var_item.get('spec')})，但在 Control Plan 之 SC 工序中未明確載明 `{g_name}`。"
                )

    # 4.1 Check Drawing SC Diamond Flag Notes coverage in Control Plan
    if dwg_sc_notes:
        for sc in dwg_sc_notes:
            note_num = sc["note_num"]
            content = sc["content"]
            matched_sc = False
            matching_rows = []
            
            for row in cp_rows:
                is_sc_class = any(k in row["special_char"].upper() for k in ['SC', 'CTQ', 'SIGNIFICANT', 'CRITICAL'])
                txt_all = f"{row['proc_name']} {row['prod_char']} {row['proc_char']} {row['spec']}".lower()
                
                is_match = False
                if "material" in content.lower() or "paper" in content.lower() or "substrate" in content.lower():
                    if any(k in txt_all for k in ["incoming", "material", "raw", "paper", "substrate", "roll", "iqc"]) or (is_sc_class and any(k in txt_all for k in ["printing", "appearance", "label", "spec"])):
                        is_match = True
                elif any(k in content.lower() for k in ["perforated", "die cut", "style", "cut", "score"]):
                    if any(k in txt_all for k in ["printing", "die cut", "cut", "perforat", "press", "slitting", "forming", "appearance"]):
                        is_match = True
                        
                if is_match:
                    matching_rows.append(row)
                    if is_sc_class:
                        matched_sc = True
                        break
                        
            if matched_sc:
                continue
            elif matching_rows:
                results["status"] = "FAIL"
                first_row = matching_rows[0]
                results["findings"].append(
                    f"❌ **FAIL (圖面 SC 缺漏)**: 圖面附註 Note {note_num} (`◇ {content}`) 標註為 SC 重點特性，"
                    f"但 Control Plan 工序 {first_row['proc_num']} [{first_row['proc_name']}] (Row {first_row['row']}) 的分類欄 (Classification) 填寫為 `{first_row['special_char'] or '未填'}`，未宣告為 SC！"
                )
            else:
                results["status"] = "FAIL"
                results["findings"].append(
                    f"❌ **FAIL (圖面 SC 未納管)**: 圖面附註 Note {note_num} (`◇ {content}`) 標註為 SC 重點特性，但在 Control Plan 中完全未找到對應管制工序。"
                )
                
    # 5. Check Drawing process flag notes in Control Plan (e.g. Paint thickness)
    if dwg_specs and "text" in dwg_specs:
        dwg_txt = dwg_specs["text"]
        m_paint = re.search(r'(\d+)\s*(?:±|#|\+/-)\s*(\d+)\s*(?:MICRON|UM|µM)', dwg_txt, re.IGNORECASE)
        if m_paint:
            paint_spec_str = f"{m_paint.group(1)} ± {m_paint.group(2)} Micron"
            paint_rows = [r for r in cp_rows if "paint" in r["proc_name"].lower() or "spray" in r["proc_name"].lower()]
            for pr in paint_rows:
                if m_paint.group(1) not in pr["spec"]:
                    results["findings"].append(
                        f"⚠️ **工程附註製程覆蓋提醒**：圖面附註規範噴漆厚度為 `{paint_spec_str}`，但 Control Plan 工序 {pr['proc_num']} [{pr['proc_name']}] (Row {pr['row']}) 之規格欄僅填寫 `{pr['spec']}`，建議補上漆厚數值管制。"
                    )

    # 5.1 Check Control Plan Reaction Plan Statistical Robustness for Sampling Plans
    # Requirement: Discarding failed units on a sampling plan is not a robust statistical control;
    # at minimum, should be sorting units that were produced since the last inspection.
    for row in cp_rows:
        sample_str = str(row.get("sample_size", "")).lower()
        freq_str = str(row.get("freq", "")).lower()
        reaction_str = str(row.get("reaction_plan", "")).strip()
        proc_name_lower = str(row.get("proc_name", "")).lower()
        prod_char_lower = str(row.get("prod_char", "")).lower()
        
        # Check if row is on a sampling plan (not 100% inspection, and not incoming cert review)
        is_100_percent = any(k in sample_str or k in freq_str for k in ["100%", "all units", "full inspection"])
        is_incoming_cert = any(k in proc_name_lower or k in prod_char_lower for k in ["incoming", "raw material"]) and "cert" in f"{row.get('spec', '')} {row.get('eval_tech', '')}".lower()
        
        has_sample_indication = bool(re.search(r'\b\d+\b', sample_str) or any(k in sample_str for k in ["pcs", "sample", "ea", "pc"]))
        has_freq_indication = any(k in freq_str for k in ["hour", "hr", "shift", "lot", "day", "min", "per", "every"])
        
        is_sampling_plan = (has_sample_indication or has_freq_indication) and not is_100_percent and not is_incoming_cert
        
        if is_sampling_plan:
            if not reaction_str:
                results["status"] = "FAIL"
                results["findings"].append(
                    f"❌ **CRITICAL REACTION PLAN MISSING (Reaction Plan 處置計畫空白)**: Control Plan 工序 {row['proc_num']} [{row['proc_name']}] - {row['prod_char']} (Row {row['row']}) 為抽樣管制項目 (Sample: {row['sample_size']}, Freq: {row['freq']})，但 Reaction Plan 欄位完全空白！當抽樣不合格時缺乏停機、隔離與回溯挑檢自前次檢驗合格以來產品 (Sort units produced since last inspection) 之圍堵措施。"
                )
            else:
                rx_lower = reaction_str.lower()
                has_sorting = any(k in rx_lower for k in ["sort", "sorting", "100% inspection", "100% check", "screen", "screening", "挑檢", "全檢", "回溯"])
                has_since_last_insp = any(k in rx_lower for k in [
                    "since last", "since the last", "previous inspection", "last inspection", "last check",
                    "last good", "前次檢驗", "上次檢驗", "自前次", "前批"
                ])
                
                is_merely_discard = any(k in rx_lower for k in ["discard", "scrap", "報廢", "丟棄"]) and not (has_sorting or has_since_last_insp)
                
                if is_merely_discard:
                    results["status"] = "FAIL"
                    results["findings"].append(
                        f"❌ **CRITICAL REACTION PLAN DEFICIENCY (抽樣計畫異常處置欠缺統計管制 / 漏列回溯挑檢)**: Control Plan 工序 {row['proc_num']} [{row['proc_name']}] - {row['prod_char']} (Row {row['row']}) 為抽樣檢驗計畫 (Sample: {row['sample_size']}, Freq: {row['freq']})，但其 Reaction Plan 僅標註為 '{reaction_str}'。依據統計製程管制與 ASSA ABLOY PQP 審查標準，在抽樣檢驗中僅丟棄抽樣失敗品 (Discarding failed units on a sampling plan) 不具備穩健的統計管制能力 (not a robust statistical control)；當抽樣不合格時，處置計畫至少必須包含回溯挑檢自前次檢驗合格以來所生產的全部產品 (At minimum, should be sorting units that were produced since the last inspection) 並對可疑品執行隔離與重工/報廢。"
                    )
                elif not (has_sorting and has_since_last_insp) and not any(k in rx_lower for k in ["return to supplier", "repacking", "quarantine"]):
                    results["findings"].append(
                        f"⚠️ **REACTION PLAN CONTAINMENT WARNING (異常處置建議具體載明回溯挑檢)**: Control Plan 工序 {row['proc_num']} [{row['proc_name']}] - {row['prod_char']} (Row {row['row']}) 為抽樣檢驗項目 (Sample: {row['sample_size']}, Freq: {row['freq']})，目前處置為 '{reaction_str}'。建議明確補充「回溯挑檢自前次檢驗合格以來之產品」(Sort units produced since last inspection)，以符合嚴格的統計製程管制要求。"
                    )

    results["sc_ctq_dict"] = dim_sc_items
    results["all_dim_balloons"] = all_dim_balloons
    return results

def run_inspection_plan_audit(wb, dwg_specs=None, dim_sc_items=None, all_dim_balloons=None):
    """
    Audits coverage, traceability, and specificity of dimensional balloons and
    Significant Characteristics (SC / CTQ) in the Inspection Plan worksheet, cross-referencing against:
    - 100% Dimensional Results (all balloons #1..N and SC/CTQ declared items)
    - Drawing notes (material, finish, appearance)
    """
    results = {
        "status": "PASS",
        "inspection_sheet": None,
        "total_dim_balloons": 0,
        "total_sc_balloons": 0,
        "sc_balloons_declared": [],
        "sc_balloons_in_insp": [],
        "sc_balloons_missing_in_insp": [],
        "all_balloons_in_insp": [],
        "all_balloons_missing_in_insp": [],
        "generic_entries": [],
        "inspection_rows": [],
        "findings": []
    }

    # 1. Locate Inspection Plan worksheet
    insp_sheets = [s for s in wb.sheetnames if "inspection plan" in s.lower() and wb[s].sheet_state == 'visible']
    if not insp_sheets:
        results["status"] = "WARNING"
        results["findings"].append("⚠️ **未找到檢驗計畫分頁**：可見工作表中未找到 [Inspection Plan] 工作表。")
        return results

    results["inspection_sheet"] = insp_sheets[0]
    ws_insp = wb[insp_sheets[0]]

    # 2. Extract dim balloons if not provided
    if all_dim_balloons is None or dim_sc_items is None:
        dim_sheets = [s for s in wb.sheetnames if "Dimensional" in s and wb[s].sheet_state == 'visible']
        all_dim_balloons = {}
        dim_sc_items = {}
        if dim_sheets:
            ws_dim = wb[dim_sheets[0]]
            cols_dim = find_dimensional_columns(ws_dim)
            sc_c_idx = 3
            for c in range(1, 15):
                val = str(ws_dim.cell(row=24, column=c).value or '').upper()
                if 'SC' in val or 'CTQ' in val or 'SIGNIFICANT' in val:
                    sc_c_idx = c
                    break
            for r in range(25, ws_dim.max_row + 1):
                b_val = ws_dim.cell(row=r, column=cols_dim.get("balloon", 2)).value
                sc_val = ws_dim.cell(row=r, column=sc_c_idx).value
                nom = ws_dim.cell(row=r, column=cols_dim.get("nominal", 4)).value
                up = ws_dim.cell(row=r, column=cols_dim.get("tol_up", 5)).value
                lo = ws_dim.cell(row=r, column=cols_dim.get("tol_lo", 6)).value
                tool = ws_dim.cell(row=r, column=13).value
                unit = ws_dim.cell(row=r, column=cols_dim.get("unit", 14)).value
                if b_val is not None:
                    try:
                        b_num = int(float(str(b_val).strip()))
                    except Exception:
                        b_num = b_val
                    is_sc = bool(sc_val and any(k in str(sc_val).upper() for k in ['SC', 'CTQ', 'SIGNIFICANT', 'CRITICAL']))
                    if b_num not in all_dim_balloons:
                        all_dim_balloons[b_num] = {
                            "row": r,
                            "nominal": nom,
                            "tol_up": float(up) if isinstance(up, (int, float)) else 0.0,
                            "tol_lo": float(lo) if isinstance(lo, (int, float)) else 0.0,
                            "sc_val": str(sc_val or '').strip(),
                            "is_sc": is_sc,
                            "tool": tool,
                            "unit": unit or "mm"
                        }
                    elif is_sc:
                        all_dim_balloons[b_num]["is_sc"] = True
                    if is_sc and b_num not in dim_sc_items:
                        dim_sc_items[b_num] = all_dim_balloons[b_num]

    results["total_dim_balloons"] = len(all_dim_balloons)
    results["total_sc_balloons"] = len(dim_sc_items)
    results["sc_balloons_declared"] = sorted(list(dim_sc_items.keys()), key=lambda x: (isinstance(x, str), x))

    # 3. Dynamically find Inspection Plan columns
    col_step = 1
    col_desc = 4
    col_method = 11
    col_std = 17
    col_sample = 25
    col_freq = 51

    for r in range(12, 18):
        for c in range(1, ws_insp.max_column + 1):
            val = str(ws_insp.cell(row=r, column=c).value or '').upper().strip()
            if 'PROCESS STEP' in val or val == 'STEP':
                col_step = c
            elif 'INSPECTION DESCRIPTION' in val or 'OPERATION DESCRIPTION' in val or val == 'DESCRIPTION':
                col_desc = c
            elif 'INSPECTION METHOD' in val or val == 'METHOD':
                col_method = c
            elif 'INSPECTION STANDARD' in val or val == 'STANDARD' or 'SPECIFICATION' in val:
                col_std = c
            elif 'SAMPLE SIZE' in val:
                col_sample = c
            elif 'TRACKING' in val or 'FREQ' in val:
                col_freq = c

    # 4. Extract Inspection Plan data rows
    insp_rows = []
    for r in range(16, ws_insp.max_row + 1):
        vals = [ws_insp.cell(row=r, column=c).value for c in range(1, ws_insp.max_column + 1) if ws_insp.cell(row=r, column=c).value is not None]
        if not vals:
            continue
        step_val = ws_insp.cell(row=r, column=col_step).value
        desc_val = str(ws_insp.cell(row=r, column=col_desc).value or '').strip()
        method_val = str(ws_insp.cell(row=r, column=col_method).value or '').strip()
        std_val = str(ws_insp.cell(row=r, column=col_std).value or '').strip()
        sample_val = str(ws_insp.cell(row=r, column=col_sample).value or '').strip()
        freq_val = str(ws_insp.cell(row=r, column=col_freq).value or '').strip()

        row_text = f"{step_val} {desc_val} {method_val} {std_val} {sample_val} {freq_val}"

        insp_rows.append({
            "row": r,
            "step": step_val,
            "desc": desc_val,
            "method": method_val,
            "std": std_val,
            "sample": sample_val,
            "freq": freq_val,
            "row_text": row_text
        })
    results["inspection_rows"] = insp_rows

    # 5. Search for balloon occurrences in Inspection Plan
    detected_balloons = set()
    for row_info in insp_rows:
        txt = row_info["row_text"]
        # Match standalone or prefixed balloon numbers (#24, Balloon 24, Balloon #24, No. 24, B24)
        matches = re.findall(r'(?:#|balloon\s*#?|no\.\s*|b\b\s*)(\d+)', txt, re.IGNORECASE)
        for m in matches:
            try:
                b_int = int(m)
                if b_int in all_dim_balloons:
                    detected_balloons.add(b_int)
            except ValueError:
                pass

        # Match ranges like Balloon #1 ~ #12 or 1-62
        range_matches = re.findall(r'(?:balloon|#|no\.)?\s*#?(\d+)\s*(?:~|-|to)\s*#?(\d+)', txt, re.IGNORECASE)
        for start_b, end_b in range_matches:
            try:
                s = int(start_b)
                e = int(end_b)
                if s <= e and (e - s) < 200:
                    for b in range(s, e + 1):
                        if b in all_dim_balloons:
                            detected_balloons.add(b)
            except ValueError:
                pass

        # Also detect generic dimension mentions with blank standard
        desc_lower = row_info["desc"].lower()
        if any(w in desc_lower for w in ["dimension", "dimensions", "size", "measurement", "caliper"]):
            if not row_info["std"] or row_info["std"].lower() in ["none", "n/a", "-", ""]:
                results["generic_entries"].append({
                    "row": row_info["row"],
                    "step": row_info["step"],
                    "desc": row_info["desc"],
                    "method": row_info["method"]
                })

    results["all_balloons_in_insp"] = sorted(list(detected_balloons))
    results["all_balloons_missing_in_insp"] = sorted([b for b in all_dim_balloons if b not in detected_balloons], key=lambda x: (isinstance(x, str), x))

    sc_detected = [b for b in results["sc_balloons_declared"] if b in detected_balloons]
    sc_missing = [b for b in results["sc_balloons_declared"] if b not in detected_balloons]
    results["sc_balloons_in_insp"] = sc_detected
    results["sc_balloons_missing_in_insp"] = sc_missing

    # 6. Formulate findings and status
    if not insp_rows:
        results["status"] = "FAIL"
        results["findings"].append("❌ **FAIL**: [Inspection Plan] 工作表無任何檢驗資料列。")
        return results

    # SC balloon coverage check
    if len(sc_missing) > 0 and results["total_sc_balloons"] > 0:
        results["status"] = "FAIL"
        sc_missing_strs = [f"Balloon #{b}" for b in sc_missing]
        results["findings"].append(
            f"❌ **CRITICAL OMISSION (SC 重點尺寸未在檢驗計畫標示)**: 全尺寸量測表已宣告 {len(dim_sc_items)} 項 SC 重點管制尺寸，"
            f"但在 [Inspection Plan] 中有 {len(sc_missing)} 項完全未列出或標註！\n"
            f"  - 未標示之 SC 球標: `{', '.join(sc_missing_strs)}`\n"
            f"  - 依品質管制規範，所有 SC/CTQ 特性必須在 Inspection Plan 明確列出其球標編號、規格公差、檢驗工具與抽樣頻率。"
        )
    elif results["total_sc_balloons"] > 0:
        results["findings"].append(
            f"✅ **PASS**: 全部 {results['total_sc_balloons']} 項 SC 重點管制尺寸均已在 [Inspection Plan] 正確標示與追溯。"
        )

    # Generic dimension entries check
    if results["generic_entries"]:
        results["status"] = "FAIL" if results["status"] == "FAIL" else "WARNING"
        gen_details = [f"Row {g['row']} (Step {g['step']}: '{g['desc']}', Method: '{g['method']}')" for g in results["generic_entries"]]
        results["findings"].append(
            f"⚠️ **檢驗標準空白與概括性規格提醒**: [Inspection Plan] 在 {', '.join(gen_details)} 填寫概括性尺寸檢驗，"
            f"但規格標準 (Inspection Standard) 欄位為空白，且完全未註明任何具體球標號碼（如 Balloon #1~#{len(all_dim_balloons)}），產線檢驗員無法得知抽樣判定基準。"
        )

    # General balloon coverage summary
    if len(detected_balloons) == 0:
        results["findings"].append(
            f"ℹ️ **全尺寸球標涵蓋統計**: 尺寸表共 {len(all_dim_balloons)} 顆球標，[Inspection Plan] 目前記載 **0 顆**具體球標（全數未展開）。建議依檢具類型（如 IM-X1000 / 游標卡尺）將球標分類展開寫入。"
        )
    else:
        coverage_pct = (len(detected_balloons) / len(all_dim_balloons)) * 100 if all_dim_balloons else 0
        results["findings"].append(
            f"ℹ️ **全尺寸球標涵蓋統計**: [Inspection Plan] 已標示 {len(detected_balloons)} / {len(all_dim_balloons)} 顆球標（涵蓋率 {coverage_pct:.1f}%）。"
        )

    # Check non-dimensional balloons (e.g. Finish / Material)
    all_insp_text = " ".join([r["row_text"].lower() for r in insp_rows])
    has_finish = any(k in all_insp_text for k in ["finish", "plate", "plating", "zinc", "chromate", "coating", "nss", "salt spray"])
    if not has_finish:
        results["findings"].append(
            "⚠️ **表面處理檢驗缺漏提醒**: 全尺寸量測表與圖面要求表面處理規格（如鍍鋅封孔 / Finish），但 [Inspection Plan] 中未找到任何表面處理（電鍍/膜厚/外觀）之檢驗工序與判定標準。"
        )

    return results

def get_sheet_drawing_texts(pqp_path):
    sheet_texts = {}
    if not pqp_path:
        return sheet_texts
    try:
        with zipfile.ZipFile(io.BytesIO(read_file_bytes_safe(pqp_path)), 'r') as z:
            wb_xml = z.read('xl/workbook.xml')
            wb_tree = ET.fromstring(wb_xml)
            sheets = {}
            for s in wb_tree.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
                r_id = s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
                s_name = s.attrib['name']
                sheets[r_id] = s_name

            wb_rels_xml = z.read('xl/_rels/workbook.xml.rels')
            wb_rels_tree = ET.fromstring(wb_rels_xml)
            sheet_targets = {}
            for rel in wb_rels_tree.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                r_id = rel.attrib['Id']
                if r_id in sheets:
                    sheet_targets[rel.attrib['Target']] = sheets[r_id]

            for ws_target, s_name in sheet_targets.items():
                ws_filename = ws_target.split('/')[-1]
                ws_rels_path = f'xl/worksheets/_rels/{ws_filename}.rels'
                if ws_rels_path in z.namelist():
                    ws_rels_tree = ET.fromstring(z.read(ws_rels_path))
                    for rel in ws_rels_tree.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                        if 'drawing' in rel.attrib.get('Type', '').lower():
                            drawing_target = rel.attrib.get('Target')
                            drawing_path = 'xl/' + drawing_target.lstrip('../') if not drawing_target.startswith('xl/') else drawing_target
                            if drawing_path in z.namelist():
                                d_xml = z.read(drawing_path).decode('utf-8', errors='ignore')
                                texts = re.findall(r'<a:t[^>]*>([^<]+)</a:t>', d_xml)
                                combined = " ".join(texts)
                                sheet_texts[s_name] = combined
    except Exception:
        pass
    return sheet_texts


def run_msa_and_capability_audit(wb, pqp_path=None, dwg_specs=None):
    """
    Audits Variable R&R, Attribute MSA, and Process Capability sheets.
    Verifies 100% data presence, #DIV/0! detection, and quality thresholds:
      - % Contribution < 9%
      - % Study Var < 20%
      - % Tolerance < 20%
      - Number of Distinct Categories (ndc) > 4
      - SCREEN % EFFECTIVE SCORE = 100%
      - SCREEN % EFFECTIVE SCORE vs. ATTRIBUTE (4) = 100%
    """
    anomalies = []
    audit_results = []
    active_sheets = [s.title for s in wb.worksheets if s.sheet_state == 'visible']
    sheet_drawing_texts = get_sheet_drawing_texts(pqp_path) if pqp_path else {}
    
    # Check if Variable R&R is actually required by the active PQS template
    pqs_sheets = [s for s in wb.sheetnames if s.endswith("PQS") and wb[s].sheet_state == 'visible']
    grr_required = True
    is_level_3 = False
    if pqs_sheets:
        pqs_ws = wb[pqs_sheets[0]]
        for r in range(1, 40):
            if pqs_ws.cell(row=r, column=3).value == 'GR&R':
                req_val = pqs_ws.cell(row=r, column=4).value
                if req_val is False or str(req_val).lower() == 'false':
                    grr_required = False
                    break
            row_txt = " ".join([str(pqs_ws.cell(row=r, column=c).value or '') for c in range(1, min(20, pqs_ws.max_column + 1))])
            if "LEVEL 3" in row_txt.upper() and ("PACKAGING" in row_txt.upper() or "PUCK" in row_txt.upper()):
                is_level_3 = True
                
    # 1. Variable R&R Audit
    child_var_sheets = [s for s in active_sheets if (s.startswith('Variable R&R (') or s.startswith('Variable R&R-'))]
    var_sheets_to_audit = child_var_sheets if child_var_sheets else ([s for s in ['Variable R&R'] if s in active_sheets])
    
    if var_sheets_to_audit and grr_required:
        for var_sheet_name in var_sheets_to_audit:
            vr_txt = sheet_drawing_texts.get(var_sheet_name, '')
            is_na_marked = any(k in vr_txt.lower() for k in ['n/a', 'no need to provide', 'not applicable', 'exempt', 'level 3'])
            
            if not is_na_marked:
                sheet = wb[var_sheet_name]
                
                # Check formula errors like #DIV/0!
                for r in range(40, min(55, sheet.max_row + 1)):
                    for c in [8, 10, 12]:
                        val = sheet.cell(row=r, column=c).value
                        if val is not None:
                            val_s = str(val).strip()
                            if val_s.startswith("#"):
                                anomalies.append({
                                    "sheet": var_sheet_name,
                                    "metric": f"Row {r} Col {c} Formula Error",
                                    "reason": f"量測數據未填寫或計算錯誤導致公式異常: '{val_s}'"
                                })
                                break
                    if anomalies and anomalies[-1].get("sheet") == var_sheet_name and "Formula Error" in anomalies[-1].get("metric", ""):
                        break

                # Find Total Gage R&R row
                total_g_row = None
                for r in range(1, 100):
                    val = sheet.cell(row=r, column=2).value
                    if val and str(val).strip() == 'Total Gage R&R':
                        total_g_row = r
                        break
                        
                # Find Number of Distinct Categories row
                ndc_row = None
                ndc_col = None
                for r in range(1, 100):
                    for c in range(1, 10):
                        val = sheet.cell(row=r, column=c).value
                        if val and "Number of Distinct Categories" in str(val):
                            ndc_row = r
                            ndc_col = c
                            break
                    if ndc_row:
                        break
                        
                if total_g_row:
                    col_metric_map = {
                        8: "% Contribution",
                        10: "% Study Var",
                        12: "% Tolerance"
                    }
                    for col, metric in col_metric_map.items():
                        cell_obj = sheet.cell(row=total_g_row, column=col)
                        try:
                            val_float = parse_float_val(cell_obj, metric)
                            audit_results.append({
                                "sheet": var_sheet_name,
                                "metric": metric,
                                "row": total_g_row,
                                "col": col,
                                "value": val_float
                            })
                        except ValueError as e:
                            anomalies.append({
                                "sheet": var_sheet_name,
                                "metric": metric,
                                "row": total_g_row,
                                "col": col,
                                "reason": str(e)
                            })
                else:
                    anomalies.append({
                        "sheet": var_sheet_name,
                        "metric": "Total Gage R&R Row",
                        "row": None,
                        "col": None,
                        "reason": "找不到 'Total Gage R&R' 數據列"
                    })
                    
                if ndc_row:
                    ndc_cell = None
                    for offset in range(1, 10):
                        c_obj = sheet.cell(row=ndc_row, column=ndc_col + offset)
                        if c_obj.value is not None:
                            ndc_cell = c_obj
                            break
                    try:
                        val_float = parse_float_val(ndc_cell, "Number of Distinct Categories")
                        audit_results.append({
                            "sheet": var_sheet_name,
                            "metric": "Number of Distinct Categories",
                            "row": ndc_row,
                            "col": ndc_cell.column if ndc_cell else ndc_col + offset,
                            "value": val_float
                        })
                    except ValueError as e:
                        anomalies.append({
                            "sheet": var_sheet_name,
                            "metric": "Number of Distinct Categories",
                            "row": ndc_row,
                            "col": ndc_cell.column if ndc_cell else ndc_col,
                            "reason": str(e)
                        })
                else:
                    anomalies.append({
                        "sheet": var_sheet_name,
                        "metric": "Number of Distinct Categories Row",
                        "row": None,
                        "col": None,
                        "reason": "找不到 'Number of Distinct Categories' 數據列"
                    })
                    
    elif not var_sheets_to_audit and grr_required and not is_level_3:
        if pqs_sheets:
            anomalies.append({
                "sheet": "[Variable R&R]",
                "metric": "Sheet Existence",
                "reason": "缺少必要之 'Variable R&R' 工作表，且 PQS 要求提供"
            })
            
    # 2. Attribute R&R Audit
    attribute_sheets = [s for s in active_sheets if "Attribute R&R" in s]
    for sname in attribute_sheets:
        sheet = wb[sname]
        
        # Check if Attribute R&R is actually populated or marked N/A
        has_trials = False
        for r in range(14, min(44, sheet.max_row + 1)):
            for c in [4, 5, 7, 8, 10, 11]:
                if sheet.cell(row=r, column=c).value is not None:
                    has_trials = True
                    break
            if has_trials:
                break
        if not has_trials:
            continue
        
        eff_row = None
        eff_col = None
        for r in range(1, 100):
            for c in range(1, 20):
                val = sheet.cell(row=r, column=c).value
                if val and "SCREEN % EFFECTIVE SCORE" in str(val) and "ATTRIBUTE (4)" not in str(val):
                    eff_row = r
                    eff_col = c
                    break
            if eff_row:
                break
                
        attr4_row = None
        attr4_col = None
        for r in range(1, 100):
            for c in range(1, 20):
                val = sheet.cell(row=r, column=c).value
                if val and "SCREEN % EFFECTIVE SCORE" in str(val) and "ATTRIBUTE (4)" in str(val):
                    attr4_row = r
                    attr4_col = c
                    break
            if attr4_row:
                break
                
        if eff_row:
            eff_cell = None
            for offset in range(1, 10):
                c_obj = sheet.cell(row=eff_row, column=eff_col + offset)
                if c_obj.value is not None:
                    eff_cell = c_obj
                    break
            try:
                val_float = parse_float_val(eff_cell, "SCREEN % EFFECTIVE SCORE")
                audit_results.append({
                    "sheet": sname,
                    "metric": "SCREEN % EFFECTIVE SCORE",
                    "row": eff_row,
                    "col": eff_cell.column if eff_cell else eff_col + offset,
                    "value": val_float
                })
            except ValueError as e:
                anomalies.append({
                    "sheet": sname,
                    "metric": "SCREEN % EFFECTIVE SCORE",
                    "row": eff_row,
                    "col": eff_cell.column if eff_cell else eff_col,
                    "reason": str(e)
                })
                
        if attr4_row:
            attr4_cell = None
            for offset in range(1, 10):
                c_obj = sheet.cell(row=attr4_row, column=attr4_col + offset)
                if c_obj.value is not None:
                    attr4_cell = c_obj
                    break
            try:
                val_float = parse_float_val(attr4_cell, "ATTRIBUTE (4)")
                audit_results.append({
                    "sheet": sname,
                    "metric": "ATTRIBUTE (4)",
                    "row": attr4_row,
                    "col": attr4_cell.column if attr4_cell else attr4_col + offset,
                    "value": val_float
                })
            except ValueError as e:
                anomalies.append({
                    "sheet": sname,
                    "metric": "ATTRIBUTE (4)",
                    "row": attr4_row,
                    "col": attr4_cell.column if attr4_cell else attr4_col,
                    "reason": str(e)
                })

    # Validate thresholds
    for res in audit_results:
        metric = res["metric"]
        val = res["value"]
        passed = True
        reason = ""
        
        if metric == "SCREEN % EFFECTIVE SCORE":
            if val != 100.0:
                passed = False
                reason = f"實測值為 {val:.2f}%, 未達 100% 門檻"
        elif metric == "ATTRIBUTE (4)":
            if val != 100.0:
                passed = False
                reason = f"實測值為 {val:.2f}%, 未達 100% 門檻"
        elif metric == "% Contribution":
            if val >= 9.0:
                passed = False
                reason = f"實測值為 {val:.2f}%, 超出 < 9% 標準門檻"
        elif metric == "% Study Var":
            if val >= 20.0:
                passed = False
                reason = f"實測值為 {val:.2f}%, 超出 < 20% 標準門檻"
        elif metric == "% Tolerance":
            if val >= 20.0:
                passed = False
                reason = f"實測值為 {val:.2f}%, 超出 < 20% 標準門檻"
        elif metric == "Number of Distinct Categories":
            if val <= 4.0:
                passed = False
                reason = f"實測值為 {int(val)}, 未滿足 > 4 標準門檻"
                
        if not passed:
            anomalies.append({
                "sheet": res["sheet"],
                "metric": metric,
                "row": res["row"],
                "col": res["col"],
                "reason": reason
            })

    # 3. Process Capability Audit (Mandatory if CTQ characteristics exist)
    has_ctq = False
    ctq_features_desc = []
    if dwg_specs:
        for f in dwg_specs.get("sc_features", []):
            if f.get("type") == "CTQ":
                has_ctq = True
                b_str = f"Balloon #{f.get('matched_balloon')}" if f.get('matched_balloon') else f"特徵代號 {f.get('letter')}"
                ctq_features_desc.append(f"{b_str} ({f.get('nominal')}\")")

    # Check PQP sheets for CTQ declarations
    for s in active_sheets:
        if "Dimensional" in s:
            ws_d = wb[s]
            for r in range(25, min(150, ws_d.max_row + 1)):
                for c in range(1, 10):
                    val = str(ws_d.cell(row=r, column=c).value or '').upper()
                    if 'CTQ' in val:
                        has_ctq = True
                        b_val = ws_d.cell(row=r, column=2).value
                        nom_val = ws_d.cell(row=r, column=4).value
                        ctq_features_desc.append(f"Balloon #{b_val} ({nom_val}\")")
                        break

    if has_ctq:
        # 1. Index Sheet: Process Capability must NOT be marked N/A
        if 'Index' in wb.sheetnames:
            ws_idx = wb['Index']
            for r in range(10, 30):
                txt = " ".join([str(ws_idx.cell(row=r, column=c).value or '') for c in range(1, 5)])
                if 'Process Capability' in txt:
                    status_val = str(ws_idx.cell(row=r, column=12).value or '').strip()
                    if status_val.upper() in ['N/A', 'NA', 'NO', 'FALSE']:
                        anomalies.append({
                            "sheet": "Index",
                            "metric": "Process Capability Checklist",
                            "row": r,
                            "col": 12,
                            "reason": f"❌ **FAIL**: 圖面/規格已定義 CTQ 關鍵品質特性 ({', '.join(sorted(list(set(ctq_features_desc))))})，依 HHI QP-04-HHI-001 與 Note 7 強制要求 Cpk >= 1.33，但 [Index] 工作表第 {r} 列卻將 Process Capability 標記為 '{status_val}'！"
                        })
                    break

        # 2. Process Capability sheet must be populated and Cpk >= 1.33
        if 'Process Capability' in active_sheets:
            sheet = wb['Process Capability']
            empty = True
            for r in range(33, 50):
                if sheet.cell(row=r, column=3).value is not None:
                    empty = False
                    break
            if empty:
                anomalies.append({
                    "sheet": "Process Capability",
                    "metric": "Process Capability Data",
                    "row": 33,
                    "col": 3,
                    "reason": f"❌ **FAIL**: 本件含有 CTQ 關鍵品質特性 ({', '.join(sorted(list(set(ctq_features_desc))))})，依 HHI QP-04-HHI-001 與圖面 Note 7 規範，必須提供連續抽樣量測數據並達成 Cpk >= 1.33！目前活頁簿內製程能力數據完全空白。"
                })
    elif 'Process Capability' in active_sheets:
        pc_txt = sheet_drawing_texts.get('Process Capability', '')
        is_na_marked = any(k in pc_txt.lower() for k in ['n/a', 'no ctq dimension', 'no ctq', 'no need to provide', 'not applicable', 'exempt', 'level 3'])
        
        if not is_na_marked:
            sheet = wb['Process Capability']
            empty = True
            for r in range(33, 40):
                if sheet.cell(row=r, column=3).value is not None:
                    empty = False
                    break
            if empty:
                anomalies.append({
                    "sheet": "Process Capability",
                    "metric": "Sample Data",
                    "reason": "製程能力工作表無抽樣量測數據，呈空白狀態"
                })
            
    return anomalies

def run_dimensional_audit(wb):
    warnings = []
    dim_sheets = [s for s in wb.sheetnames if "Dimensional" in s and wb[s].sheet_state == 'visible']
    
    for sname in dim_sheets:
        sheet = wb[sname]
        cols = find_dimensional_columns(sheet)
        
        balloon_samples = {}
        for r in range(25, sheet.max_row + 1):
            balloon = sheet.cell(row=r, column=cols["balloon"]).value
            nom = sheet.cell(row=r, column=cols["nominal"]).value
            usl = sheet.cell(row=r, column=cols.get("tol_up", 5) + 2 if cols.get("tol_up") else 7).value
            lsl = sheet.cell(row=r, column=cols.get("tol_lo", 6) + 2 if cols.get("tol_lo") else 8).value
            meas = sheet.cell(row=r, column=cols["meas"]).value
            disp = sheet.cell(row=r, column=cols["disp"]).value
            sample = sheet.cell(row=r, column=cols["sample"]).value
            
            if balloon is not None and sample is not None:
                try:
                    b_num = int(float(str(balloon).strip()))
                    s_num = int(float(str(sample).strip()))
                    balloon_samples.setdefault(b_num, set()).add(s_num)
                except (ValueError, TypeError):
                    pass
            
            if balloon is not None and nom is not None and disp is not None:
                disp_str = str(disp).strip()
                # Check for formula tampering: value is out of spec, but disp is hardcoded 'Acceptable'
                is_tampered_pass = False
                try:
                    m_f = float(str(meas).strip())
                    u_f = float(str(usl).strip())
                    l_f = float(str(lsl).strip())
                    if (m_f > u_f + 1e-5 or m_f < l_f - 1e-5) and disp_str.lower() == 'acceptable':
                        is_tampered_pass = True
                except Exception:
                    pass

                if is_tampered_pass:
                    warnings.append({
                        "sheet": sname,
                        "row": r,
                        "balloon": balloon,
                        "nominal": nom,
                        "measured": meas,
                        "disposition": f"❌ **CRITICAL FAIL (公式遭覆寫掩蓋超差)**: 實測值 {meas} 超出規格上下限 [{lsl}, {usl}]，但判定欄公式遭覆寫為固定字串 'Acceptable'！(註：若此項為 GD&T 單邊公差或 MAX R，請將下差修正為 -{nom} 使 LSL = 0，並還原判定欄公式)"
                    })
                elif disp_str.lower() != 'acceptable' and disp_str != ' ' and disp_str != '':
                    warnings.append({
                        "sheet": sname,
                        "row": r,
                        "balloon": balloon,
                        "nominal": nom,
                        "measured": meas,
                        "disposition": disp_str
                    })
                    
        # Check sample count completeness
        if balloon_samples:
            max_s = max(len(s) for s in balloon_samples.values())
            if max_s < 5:
                warnings.append({
                    "sheet": sname,
                    "row": "Header",
                    "balloon": "All",
                    "nominal": "全項目",
                    "measured": f"{max_s} pcs",
                    "disposition": f"抽樣數量提醒：各尺寸僅量測 {max_s} 顆樣品 (Sample #1~#{max_s})。ASSA ABLOY 標準通常要求提供 5 顆 (Sample #1~#5)，請確認本案是否特許 3 顆樣品"
                })
                    
    return warnings, dim_sheets

def run_dimensional_spec_cross_check(wb, dwg_specs):
    """
    Cross-checks the dimensions in the PQP workbook's Dimensional sheets against the PDF drawing specifications.
    """
    results = []
    dim_sheets = [s for s in wb.sheetnames if "Dimensional" in s and wb[s].sheet_state == 'visible']
    
    if not dwg_specs or "dimensions" not in dwg_specs:
        return [{"status": "INFO", "message": "無圖面尺寸規格可供交叉比對（未提供圖面或解析失敗）"}]
        
    pdf_dims = dwg_specs["dimensions"]
    gen_tol = dwg_specs["general_tolerance"]
    pdf_text = dwg_specs["text"]
    
    for sname in dim_sheets:
        sheet = wb[sname]
        cols = find_dimensional_columns(sheet)
        excel_balloons = {}
        
        # 1. Collect unique balloon entries from Excel
        for r in range(26, sheet.max_row + 1):
            cell_b = sheet.cell(row=r, column=cols["balloon"])
            balloon = cell_b.value
            cell_nom = sheet.cell(row=r, column=cols["nominal"])
            nominal = cell_nom.value
            cell_tu = sheet.cell(row=r, column=cols["tol_up"])
            tol_up = cell_tu.value
            cell_tl = sheet.cell(row=r, column=cols["tol_lo"])
            tol_lo = cell_tl.value
            unit = sheet.cell(row=r, column=cols["unit"]).value
            
            if balloon is not None:
                balloon_id = str(balloon).strip()
                if balloon_id not in excel_balloons:
                    excel_balloons[balloon_id] = {
                        "rows": [r],
                        "nominal": nominal,
                        "tol_up": tol_up,
                        "tol_lo": tol_lo,
                        "unit": unit,
                        "nom_fmt": getattr(cell_nom, 'number_format', None),
                        "tol_up_fmt": getattr(cell_tu, 'number_format', None),
                        "tol_lo_fmt": getattr(cell_tl, 'number_format', None)
                    }
                else:
                    excel_balloons[balloon_id]["rows"].append(r)
        
        # 2. Compare each balloon with drawing
        for b_id, entry in excel_balloons.items():
            nominal = entry["nominal"]
            tol_up = entry["tol_up"]
            tol_lo = entry["tol_lo"]
            unit = entry["unit"]
            
            # Normalize nominal: extract numeric part and suffix (e.g. 0.625(MIN) -> 0.625, MIN)
            clean_nom = nominal
            suffix = ""
            if isinstance(nominal, str):
                nominal_str = nominal.strip()
                m_suffix = re.match(r'^([\d\.\-\/]+)\s*\(?(MIN|MAX|REF|MINIMUM|MAXIMUM)\)?$', nominal_str, re.IGNORECASE)
                if m_suffix:
                    clean_nom = m_suffix.group(1)
                    suffix = m_suffix.group(2).upper()
            
            is_numeric = False
            nom_val = None
            if isinstance(clean_nom, (int, float)):
                is_numeric = True
                nom_val = float(clean_nom)
            elif isinstance(clean_nom, str):
                clean_nom_str = clean_nom.strip()
                if '/' in clean_nom_str:
                    try:
                        nom_val = parse_fraction(clean_nom_str)
                        is_numeric = True
                    except Exception:
                        pass
                elif re.match(r'^\d*(\.\d+)?$', clean_nom_str) and clean_nom_str not in ['', '.']:
                    nom_val = float(clean_nom_str)
                    is_numeric = True
            
            if is_numeric:
                is_mm = False
                is_angle = False
                if unit is not None:
                    unit_str = str(unit).lower()
                    if "mm" in unit_str:
                        is_mm = True
                    elif "in" in unit_str or "inch" in unit_str:
                        is_mm = False
                    elif "deg" in unit_str or "°" in unit_str or "x" in unit_str:
                        is_angle = True
                    elif nom_val > 5.0:
                        is_mm = True
                elif nom_val > 5.0:
                    is_mm = True
                
                # Find the closest nominal from the PDF
                closest_dim = None
                min_diff = 999.0
                candidates = []
                
                for pd in pdf_dims:
                    pd_is_angle = pd.get("is_angle", False) or "°" in str(pd.get("inch_str", "")) or "$" in str(pd.get("inch_str", ""))
                    if is_angle:
                        if not pd_is_angle:
                            continue
                        diff = abs(pd.get("inch_val", 0.0) - nom_val)
                    elif is_mm:
                        if pd_is_angle:
                            continue
                        diff = abs(pd.get("mm_val", 0.0) - nom_val)
                    else:
                        if pd_is_angle:
                            continue
                        diff = abs(pd.get("inch_val", 0.0) - nom_val)
                    
                    if diff < min_diff - 1e-6:
                        min_diff = diff
                        candidates = [pd]
                    elif abs(diff - min_diff) <= 1e-6:
                        candidates.append(pd)
                
                if candidates:
                    # Break ties by matching implied decimal precision/tolerance
                    tol_u_f = float(tol_up) if tol_up is not None else None
                    if tol_u_f is not None:
                        best_match = candidates[0]
                        for c_cand in candidates:
                            c_str = c_cand.get("inch_str", "")
                            m_d = re.search(r'\.(\d+)', c_str)
                            c_dec = len(m_d.group(1)) if m_d else None
                            if c_dec == 2 and abs(tol_u_f - 0.01) < 1e-4:
                                best_match = c_cand
                                break
                            elif c_dec == 3 and abs(tol_u_f - 0.005) < 1e-4:
                                best_match = c_cand
                                break
                        closest_dim = best_match
                    else:
                        closest_dim = candidates[0]
                
                dist_limit = 1.0 if is_angle else ((0.1 if nom_val < 2.0 else 1.27) if is_mm else 0.05)
                if closest_dim and min_diff <= dist_limit:
                    # Found a match!
                    if is_angle:
                        target_tol_up = closest_dim.get("inch_tol_up")
                        target_tol_lo = closest_dim.get("inch_tol_lo")
                        if target_tol_up is None or target_tol_lo is None:
                            ang_tol = dwg_specs.get("angle_tolerance", 1.0)
                            target_tol_up = ang_tol
                            target_tol_lo = ang_tol
                        tol_up_val = float(tol_up) if tol_up is not None else 0.0
                        tol_lo_val = float(tol_lo) if tol_lo is not None else 0.0
                        tol_diff_up = abs(tol_up_val - target_tol_up)
                        tol_diff_lo = abs(tol_lo_val - target_tol_lo)
                        
                        if tol_diff_up > 0.5 or tol_diff_lo > 0.5:
                            results.append({
                                "status": "FAIL",
                                "balloon": b_id,
                                "sheet": sname,
                                "message": f"角度標記 {b_id} 的公差 ({tol_up_val}/{tol_lo_val} °) 與圖面角度公差 +{target_tol_up}/-{target_tol_lo} ° 不符。"
                            })
                        else:
                            results.append({
                                "status": "PASS",
                                "balloon": b_id,
                                "sheet": sname,
                                "message": f"Nominal: {nom_val} ° (圖面: {closest_dim['inch_str']}), 公差: +{tol_up_val}/-{tol_lo_val} (圖面: +{target_tol_up}/-{target_tol_lo} °)"
                            })
                    elif tol_up is None and tol_lo is None:
                        # Basic dimension or reference dimension without explicit tolerances
                        results.append({
                            "status": "PASS",
                            "balloon": b_id,
                            "sheet": sname,
                            "message": f"Nominal: {nom_val} in (圖面: {closest_dim['inch_str']}), 基本尺寸 (Basic Dimension) 無須公差。"
                        })
                    elif suffix in ["MIN", "MAX", "REF", "MINIMUM", "MAXIMUM"]:
                        dwg_str = (closest_dim.get("inch_str") if not is_mm else closest_dim.get("mm_str")) or ""
                        m_dwg_dec = re.search(r'\.(\d+)', dwg_str)
                        req_nom_dec = len(m_dwg_dec.group(1)) if m_dwg_dec else None
                        
                        fmt_alerts = []
                        nom_fmt = entry.get("nom_fmt")
                        if req_nom_dec is not None and not is_mm:
                            nom_dec = get_fmt_decimals(nom_fmt)
                            if nom_dec is not None and nom_dec < req_nom_dec:
                                fmt_alerts.append(f"⚠️【尺寸規格小數位數不足】圖面標註為 '{dwg_str}' ({req_nom_dec} 位小數)，但儲存格格式設定為 '{nom_fmt}' ({nom_dec} 位小數)，導致末位 0 丟失！建議將尺寸規格及極限儲存格格式設為 '0.{'0'*req_nom_dec}'")
                            elif nom_fmt == 'General' and req_nom_dec >= 2:
                                fmt_alerts.append(f"⚠️【公稱值格式警告】尺寸 '{dwg_str}' 格式為 'General'，建議鎖定為固定小數位數格式以維持尾數 0")
                        
                        has_fmt_alert = len(fmt_alerts) > 0
                        extra_str = ("；" + "；".join(fmt_alerts)) if fmt_alerts else ""
                        results.append({
                            "status": "WARNING" if has_fmt_alert else "PASS",
                            "balloon": b_id,
                            "sheet": sname,
                            "message": f"Nominal: {nom_val} {suffix} (圖面: {closest_dim['inch_str']}), 此項為極限/參考尺寸，無須比對公差{extra_str}。"
                        })
                    else:
                        target_tol_up = closest_dim.get("inch_tol_up" if not is_mm else "mm_tol_up")
                        target_tol_lo = closest_dim.get("inch_tol_lo" if not is_mm else "mm_tol_lo")
                        
                        unit_label = "mm" if is_mm else "in"
                        tol_up_val = float(tol_up) if tol_up is not None else None
                        tol_lo_val = float(tol_lo) if tol_lo is not None else None
                        
                        dwg_str = (closest_dim.get("inch_str") if not is_mm else closest_dim.get("mm_str")) or ""
                        m_dwg_dec = re.search(r'\.(\d+)', dwg_str)
                        req_nom_dec = len(m_dwg_dec.group(1)) if m_dwg_dec else None

                        # Resolve general tolerance by decimal place count (ASME Y14.5M Note 4 / Title Block)
                        if target_tol_up is None and not is_mm:
                            if req_nom_dec == 2 and dwg_specs.get("tol_2_place") is not None:
                                target_tol_up = target_tol_lo = dwg_specs["tol_2_place"]
                            elif req_nom_dec == 3 and dwg_specs.get("tol_3_place") is not None:
                                target_tol_up = target_tol_lo = dwg_specs["tol_3_place"]
                            elif dwg_specs.get("general_tolerance") is not None:
                                target_tol_up = target_tol_lo = dwg_specs.get("general_tolerance")

                        # Strict Decimal Precision & Cell Format Auditing
                        fmt_alerts = []
                        tu_fmt = entry.get("tol_up_fmt")
                        tl_fmt = entry.get("tol_lo_fmt")
                        nom_fmt = entry.get("nom_fmt")
                        
                        # Never fall back to inch gen_tol (e.g. 0.005 in) when PQP entry is in mm!
                        ref_tol = target_tol_up if target_tol_up is not None else (tol_up_val if is_mm else gen_tol)
                        req_tol_dec = None
                        if ref_tol is not None:
                            s_tol = f"{ref_tol:.6f}".rstrip('0').rstrip('.')
                            if '.' in s_tol:
                                req_tol_dec = len(s_tol.split('.')[1])
                                
                        tu_dec = get_fmt_decimals(tu_fmt)
                        tl_dec = get_fmt_decimals(tl_fmt)
                        
                        # Check unilateral tolerances (tol_up == 0)
                        if tol_up_val == 0.0 and tol_lo_val is not None:
                            if abs(tol_lo_val - nom_val / 2.0) < 1e-4 and abs(nom_val - 0.01) < 1e-4:
                                fmt_alerts.append(f"⚠️【單邊公差下限設定錯誤】此項為單邊上限管制 (0 ~ {nom_val:g} in)，下差應設為 -{nom_val:g} (使下限 LSL = 0.000，與內嵌原始 FAI 一致)；目前下差誤設為 -{tol_lo_val:g} (LSL = {nom_val - tol_lo_val:g})，導致實測 < {nom_val - tol_lo_val:g} 之良品被誤判為超差！")
                        elif req_tol_dec is not None and req_tol_dec >= 3 and tu_dec is not None and tu_dec < req_tol_dec and tol_up_val is not None and tol_up_val > 0:
                            disp_rounded = round(tol_up_val, tu_dec) if tol_up_val is not None else 0.0
                            fmt_alerts.append(f"⚠️【格式四捨五入警告】公差數值為 +{tol_up_val}/-{tol_lo_val}，但儲存格格式設定為 '{tu_fmt}' ({tu_dec} 位小數)，導致 Excel 畫面上四捨五入顯示為 ±{disp_rounded}，與圖面 {req_tol_dec} 位公差 (±{ref_tol}) 不符！建議將儲存格格式改為 '0.000'")
                        elif (tu_fmt == 'General' and tol_up_val is not None and tol_up_val > 0) or (tl_fmt == 'General' and tol_lo_val is not None and tol_lo_val > 0):
                            fmt_alerts.append("⚠️【格式通用警告】公差格式為 'General'，建議設定固定小數格式 (如 '0.000') 以防尾數 0 丟失")
                            
                        if req_nom_dec is not None and not is_mm:
                            nom_dec = get_fmt_decimals(nom_fmt)
                            if nom_dec is not None and nom_dec < req_nom_dec and tol_up_val != 0.0:
                                fmt_alerts.append(f"⚠️【尺寸規格小數位數不足引發公差位數矛盾】圖面 .XXX 要求是指尺寸規格小數位數（標註為 '{dwg_str}'，{req_nom_dec} 位小數），依 Note 4 適用 ±{ref_tol} 公差；但 PQP 尺寸規格欄格式為 '{nom_fmt}' ({nom_dec} 位小數)，吃掉末位 0，造成「尺寸規格 {nom_dec} 位但公差 {req_tol_dec} 位」之位數矛盾與規格極限進位偏差！建議將尺寸規格及上限/下限儲存格格式統一設為 '0.000'")
                            elif nom_dec is not None and nom_dec > req_nom_dec:
                                fmt_alerts.append(f"⚠️【公稱值精度不一致】圖面尺寸標註為 '{dwg_str}' ({req_nom_dec} 位小數)，但儲存格格式為 '{nom_fmt}' ({nom_dec} 位小數)")
                            elif nom_fmt == 'General' and req_nom_dec >= 2:
                                fmt_alerts.append(f"⚠️【公稱值格式警告】尺寸 '{dwg_str}' 格式為 'General'，建議鎖定為 '0.000' 以維持尾數 0")
                                
                        has_rounding_bug = any("四捨五入警告" in a or "單邊公差下限設定錯誤" in a for a in fmt_alerts)
                        extra_str = ("；" + "；".join(fmt_alerts)) if fmt_alerts else ""
                        
                        if tol_up_val == 0.0 and tol_lo_val is not None:
                            # Unilateral tolerance
                            results.append({
                                "status": "WARNING" if has_rounding_bug else "PASS",
                                "balloon": b_id,
                                "sheet": sname,
                                "message": f"Nominal: {nom_val} {unit_label} (圖面單邊幾何公差 GD&T/MAX callout), 公差: +{tol_up_val}/-{tol_lo_val} {unit_label}{extra_str}"
                            })
                        elif target_tol_up is not None and target_tol_lo is not None:
                            tol_match_ok = True
                            if tol_up_val is not None and abs(tol_up_val - target_tol_up) > (1e-4 if not is_mm else 0.05):
                                tol_match_ok = False
                            if tol_lo_val is not None and abs(tol_lo_val - target_tol_lo) > (1e-4 if not is_mm else 0.05):
                                tol_match_ok = False
                                
                            if tol_match_ok:
                                results.append({
                                    "status": "WARNING" if has_rounding_bug else "PASS",
                                    "balloon": b_id,
                                    "sheet": sname,
                                    "message": f"Nominal: {nom_val} {unit_label} (圖面: {closest_dim['inch_str'] if not is_mm else closest_dim['mm_str']}), 公差: +{tol_up_val}/-{tol_lo_val} {unit_label} (圖面規格一致){extra_str}"
                                })
                            else:
                                results.append({
                                    "status": "WARNING",
                                    "balloon": b_id,
                                    "sheet": sname,
                                    "message": f"Nominal: {nom_val} {unit_label} (圖面: {closest_dim['inch_str'] if not is_mm else closest_dim['mm_str']}), 公差: +{tol_up_val}/-{tol_lo_val} (圖面標註: +{target_tol_up:.4f}/-{target_tol_lo:.4f}){extra_str}"
                                })
                        else:
                            results.append({
                                "status": "WARNING" if has_rounding_bug else "PASS",
                                "balloon": b_id,
                                "sheet": sname,
                                "message": f"Nominal: {nom_val} {unit_label} (圖面: {closest_dim['inch_str'] if not is_mm else closest_dim['mm_str']}), 公差: +{tol_up_val}/-{tol_lo_val} (通用公差適用){extra_str}"
                            })
                else:
                    tol_up_val = float(tol_up) if tol_up is not None else None
                    tol_lo_val = float(tol_lo) if tol_lo is not None else None
                    unit_label = "mm" if is_mm else "in"
                    # Handle CAD vector-drawn GD&T feature control frame callouts (e.g., Flatness .01 [0.3] -> 0.3 +0/-0.25 mm)
                    if is_mm and nom_val <= 0.5 and (tol_up_val == 0.0 or tol_lo_val == 0.0):
                        results.append({
                            "status": "PASS",
                            "balloon": b_id,
                            "sheet": sname,
                            "message": f"Nominal: {nom_val} {unit_label} (CAD 向量幾何公差框 GD&T Feature Control Frame), 公差: +{tol_up_val}/-{tol_lo_val} {unit_label}"
                        })
                    else:
                        results.append({
                            "status": "WARNING",
                            "balloon": b_id,
                            "sheet": sname,
                            "message": f"Excel 中的尺寸標記 {b_id} (Nominal: {nominal}) 在圖面 PDF 中未找到完全吻合之英吋/毫米 callout。"
                        })
            else:
                results.append({
                    "status": "INFO",
                    "balloon": b_id,
                    "sheet": sname,
                    "message": f"尺寸標記 {b_id} (Nominal: '{nominal}') 為非數值或自訂字串，略過公差比對。"
                })
                
    return results

def robust_norm(s):
    if not s:
        return ""
    s_clean = re.sub(r'[\s\-_]+', '', str(s)).upper()
    return s_clean

def parse_dwg_paths(dwg_arg):
    if not dwg_arg:
        return []
    if os.path.exists(dwg_arg.strip()):
        return [dwg_arg.strip()]
    if ',' in dwg_arg:
        parts = [p.strip() for p in dwg_arg.split(',') if p.strip()]
        if any(os.path.exists(p) for p in parts):
            return parts
    return [dwg_arg.strip()]


def run_unrelated_embeddings_audit(pqp_path, wb):
    """
    Inspects all embedded objects (OLE .bin, .xlsx, etc.) inside the PQP workbook
    and flags any file whose detected part number / drawing number does NOT match
    the PQP's own part number or drawing number.

    Detection logic:
    - XLSX embeddings: scan sharedStrings.xml + sheet values for 7-digit PN patterns.
    - OLE/PDF embeddings: locate the embedded %PDF stream, extract text via pymupdf,
      then regex-search for ASSA ABLOY-style part numbers (5xxxxxxx, 1xxxx-xxx).
    - Unrecognised binary: fall back to raw ASCII string mining.

    Returns a list of findings dicts with keys: file, size_kb, type, pns_found, status, message.
    """
    import zipfile

    # Collect PQP's own part numbers from the header sheet
    own_pns = set()
    header_sheet = (wb['TitlePg'] if 'TitlePg' in wb.sheetnames
                    else (wb['PQS'] if 'PQS' in wb.sheetnames else None))
    if header_sheet:
        for r in range(1, min(30, header_sheet.max_row + 1)):
            for c in range(1, min(30, header_sheet.max_column + 1)):
                v = str(header_sheet.cell(row=r, column=c).value or '').strip()
                # Match part numbers: 7-digit, 5-digit with dash, or 5-7 digit with dash suffix
                for m in re.findall(r'\b5\d{6}(?:-\d{3})?\b|\b\d{7}(?:-\d{3})?\b|\b\d{5}-\d{3}\b|\b1[0-9]{4}-\d{3}\b', v):
                    norm = re.sub(r'-0+$', '', m)
                    own_pns.add(norm)
                    if '-' in norm:
                        own_pns.add(norm.split('-')[0])

    # Also mine from all visible sheet header cells (row 1-15 only) for family PNs
    for sheet in wb.worksheets:
        if sheet.sheet_state != 'visible':
            continue
        for r in range(1, min(15, sheet.max_row + 1)):
            for c in range(1, min(15, sheet.max_column + 1)):
                v = str(sheet.cell(row=r, column=c).value or '').strip()
                for m in re.findall(r'\b5\d{6}(?:-\d{3})?\b|\b\d{7}(?:-\d{3})?\b|\b\d{5}-\d{3}\b|\b1[0-9]{4}-\d{3}\b', v):
                    norm = re.sub(r'-0+$', '', m)
                    own_pns.add(norm)
                    if '-' in norm:
                        own_pns.add(norm.split('-')[0])

    PN_PATTERNS = [
        r'\b5\d{6}(?:-\d{3})?\b',
        r'\b1[0-9]{4}-\d{3}\b',
        r'\b\d{7}(?:-\d{3})?\b',
        r'\b\d{5}-\d{3}\b',
    ]

    def find_pns_in_text(text):
        # Filter telephone / fax numbers (with prefixes like TEL, FAX, PHONE, 電話, 傳真, etc.)
        cleaned_text = re.sub(r'(?:TEL|FAX|PHONE|MOBILE|CELL|電話|傳真|專線|分機)[:：\s\.]*(?:\+?\d{1,4}[-\s]?)?(?:\(?\d{2,4}\)?[-\s]?)?\d{3,4}[-\s]?\d{3,4}(?:#\d+)?', ' ', text, flags=re.IGNORECASE)
        # Filter Taiwan landline patterns (02-xxxxxxx, 05-xxxxxxx, 03-xxxxxxx, etc.)
        cleaned_text = re.sub(r'\b0\d{1,2}[-–—\s]?\d{7,8}\b', ' ', cleaned_text)
        # Filter Mobile phone numbers (09xxxxxxxx)
        cleaned_text = re.sub(r'\b09\d{2}[-–—\s]?\d{3}[-–—\s]?\d{3}\b', ' ', cleaned_text)
        # Filter Tax ID (統編 / VAT / Tax ID)
        cleaned_text = re.sub(r'(?:統編|統一編號|TAX\s*ID|VAT|GUI)\s*[:：]?\s*\d{8}\b', ' ', cleaned_text, flags=re.IGNORECASE)
        # Filter Postal / Zip code
        cleaned_text = re.sub(r'(?:ZIP|POSTAL|郵遞區號)\s*[:：]?\s*\d{3,6}\b', ' ', cleaned_text, flags=re.IGNORECASE)
        # Filter out CAS registry numbers (e.g. 2156592-54-8, 79-94-7) from chemical test reports (RoHS/REACH SVHC)
        cleaned_text = re.sub(r'\b\d{2,7}-\d{2}-\d\b', ' ', cleaned_text)
        # Filter out SGS / CTI lab certificate/report tracking codes (e.g. TW/2500126/VoC, ETR25100709, SHAEC12345678, PT-2026-729L)
        cleaned_text = re.sub(r'[A-Za-z]{2,4}/\d{6,8}/[A-Za-z0-9]+', ' ', cleaned_text)
        cleaned_text = re.sub(r'\b[A-Za-z]{3}\d{2}-\d{7}(?:-\d{3,4})?(?:\.[A-Za-z0-9]+)?\b', ' ', cleaned_text)
        cleaned_text = re.sub(r'\b(?:ETR|SHAEC|CANEC|SZXEC|TAOEC|NGBEC|TW|PT-|A22)\d+[A-Za-z0-9-]*\b', ' ', cleaned_text)
        cleaned_text = re.sub(r'\b(?:REPORT|TEST|SAMPLE)\s*(?:NO|NUMBER|ID)[:：\s\.]*[A-Za-z0-9-]+', ' ', cleaned_text, flags=re.IGNORECASE)
        found = set()
        for pat in PN_PATTERNS:
            found.update(re.findall(pat, cleaned_text))
        return found

    def sniff_type(data):
        if data[:4] == b'%PDF':
            return 'pdf'
        if data[:4] == b'PK\x03\x04':
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as z:
                    nl = z.namelist()
                    if any(n.startswith('ppt/') for n in nl):
                        return 'pptx'
                    if any(n.startswith('word/') for n in nl):
                        return 'docx'
                    return 'xlsx'
            except Exception:
                return 'xlsx'
        if data[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
            return 'ole'
        return 'bin'

    def extract_text_xlsx(data):
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                texts = []
                if 'xl/sharedStrings.xml' in z.namelist():
                    xml = z.read('xl/sharedStrings.xml').decode('utf-8', errors='replace')
                    texts += re.findall(r'<t[^>]*>([^<]+)</t>', xml)
                for name in z.namelist():
                    if re.match(r'xl/worksheets/sheet\d+\.xml', name):
                        xml = z.read(name).decode('utf-8', errors='replace')
                        texts += re.findall(r'<v>([^<]+)</v>', xml)
                return ' '.join(texts[:600])
        except Exception:
            return ''

    def extract_text_pptx(data):
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                texts = []
                for name in z.namelist():
                    if name.startswith('ppt/slides/slide') and name.endswith('.xml'):
                        xml = z.read(name).decode('utf-8', errors='replace')
                        texts += re.findall(r'<a:t>([^<]+)</a:t>', xml)
                return ' '.join(texts[:600])
        except Exception:
            return ''

    def extract_text_pdf_bytes(pdf_bytes):
        try:
            import pymupdf
            doc = pymupdf.open(stream=pdf_bytes, filetype='pdf')
            return '\n'.join(page.get_text() or '' for page in doc)
        except Exception:
            return ''

    def extract_text_ole(data):
        # Try to find embedded PDF first
        pdf_idx = data.find(b'%PDF')
        if pdf_idx != -1:
            return extract_text_pdf_bytes(data[pdf_idx:])
        # Fall back to raw ASCII string mining
        strings = re.findall(rb'[ -~]{6,}', data)
        return ' '.join(s.decode('latin-1', errors='replace') for s in strings[:300])

    findings = []
    try:
        with zipfile.ZipFile(io.BytesIO(read_file_bytes_safe(pqp_path)), 'r') as zf:
            obj_to_sheets = get_embedded_to_sheets_map(zf)
            emb_names = sorted([n for n in zf.namelist() if 'xl/embeddings/' in n])
            for emb in emb_names:
                data = zf.read(emb)
                ftype = sniff_type(data)
                size_kb = len(data) / 1024
                basename = os.path.basename(emb)

                if ftype == 'xlsx':
                    text = extract_text_xlsx(data)
                elif ftype == 'pptx':
                    text = extract_text_pptx(data)
                elif ftype == 'ole':
                    text = extract_text_ole(data)
                elif ftype == 'pdf':
                    text = extract_text_pdf_bytes(data)
                else:
                    strings = re.findall(rb'[ -~]{6,}', data)
                    text = ' '.join(s.decode('latin-1', errors='replace') for s in strings[:200])

                detected_pns = find_pns_in_text(text)
                # Strip trailing -000 for comparison
                norm_detected = {re.sub(r'-0+$', '', p) for p in detected_pns}

                sheets_found = obj_to_sheets.get(basename, [])
                sheet_str = ', '.join([f'[{s}]' for s in sheets_found]) if sheets_found else '未指定分頁'

                # Determine relationship: match exact or base PN
                matched = set()
                for p in norm_detected:
                    if p in own_pns or (('-' in p) and p.split('-')[0] in own_pns):
                        matched.add(p)
                    elif any(op for op in own_pns if ('-' in op and op.split('-')[0] == p)):
                        matched.add(p)

                if not detected_pns:
                    status = 'no_pn'
                    msg = f"無法從此嵌入物件中偵測到料號（可能為圖像型 PDF 或純二進位格式）。"
                elif matched:
                    status = 'ok'
                    msg = f"偵測到的料號 {sorted(matched)} 與本 PQP 相符。"
                else:
                    status = 'unrelated'
                    msg = (f"偵測到的料號 {sorted(detected_pns)} 與本 PQP 所有已知料號均不符合，"
                           f"此嵌入檔案可能為誤植的其他產品文件。")

                findings.append({
                    'file': basename,
                    'sheets': sheets_found,
                    'sheet_str': sheet_str,
                    'path': emb,
                    'size_kb': size_kb,
                    'type': ftype,
                    'pns_found': sorted(detected_pns),
                    'status': status,
                    'message': msg,
                    'preview': text[:300].replace('\n', ' ').strip()
                })
    except Exception:
        pass

    return findings, sorted(own_pns)


def auto_fix_pqp_dimensional_formats(pqp_path, cross_check_results=None, dwg_spec_data=None):
    """
    Auto-fixes number format precision discrepancies in outer PQP dimensional worksheets
    (e.g. '100% Dimensional Results', '100% Dimensional_*').

    Identifies balloons where:
    - Drawing Note 4 requires 3 decimals (.XXX), but PQP nominal/specs were formatted as '0.00' or '0.0' or '@',
      causing visual rounding bugs (e.g. upper/lower limits showing 0.04/0.04 instead of 0.045/0.035).
    - Restores Col D (Nominal), Col E/F (Tol), Col G (Upper Spec), Col H (Lower Spec) to '0.000'.
    - Sets Col I (Measurement) to '0.0000'.
    - For MAX R radius callouts (R.010 MAX), ensures Col E, G, H are formatted as '0.00'.

    Executes via Windows Excel Win32 COM (preserving all embedded OLE objects and vector shapes)
    with an openpyxl fallback. Creates a backup file (.bak) before modifying.
    """
    import shutil

    if not os.path.exists(pqp_path):
        print(f"Error: PQP workbook not found: {pqp_path}", file=sys.stderr)
        return False

    target_3dec_balloons = set()
    target_max_r_balloons = set()

    if cross_check_results:
        for item in cross_check_results:
            b_str = str(item.get("balloon", "")).strip()
            msg = str(item.get("message", ""))
            if not b_str:
                continue
            try:
                b_val = int(b_str)
            except ValueError:
                b_val = b_str

            if "尺寸規格小數位數不足" in msg or "格式四捨五入警告" in msg or "0.000" in msg:
                # 幾何公差 (GD&T 如 Flatness .010) 或任何標註為 3 位小數 (.010) 者，必須使用 0.000 格式
                if "R.01" in msg and "MAX" in msg.upper() and ".010" not in msg:
                    target_max_r_balloons.add(b_val)
                else:
                    target_3dec_balloons.add(b_val)

    if dwg_spec_data and isinstance(dwg_spec_data, dict):
        dims = dwg_spec_data.get("dimensions", [])
        if not dims and "items" in dwg_spec_data:
            dims = dwg_spec_data.get("items", [])
        for d in dims:
            b = d.get("balloon") if "balloon" in d else d.get("id")
            t = str(d.get("text", "") or d.get("comments", ""))
            req_dec = d.get("req_nom_dec", 0) or d.get("dec", 0)
            if req_dec == 3 or re.search(r'(?<!\d)\.\d{3}(?!\d)', t) or "Flatness" in t:
                if "MAX" in t.upper() and ("R.01" in t) and not (".010" in t or req_dec == 3):
                    target_max_r_balloons.add(b)
                elif b is not None:
                    target_3dec_balloons.add(b)

    if not target_3dec_balloons and not target_max_r_balloons:
        print("[Auto-Fix] No cell format precision issues detected to fix.", flush=True)
        return True

    print(f"[Auto-Fix] Identified {len(target_3dec_balloons)} 3-decimal balloons and {len(target_max_r_balloons)} MAX R balloons to fix.", flush=True)

    bak_path = pqp_path + ".bak"
    if not os.path.exists(bak_path):
        try:
            shutil.copy2(pqp_path, bak_path)
            print(f"[Auto-Fix] Backup created: {bak_path}", flush=True)
        except Exception as e:
            print(f"[Auto-Fix] Warning: Failed to create backup: {e}", flush=True)

    repaired_count = 0
    com_success = False
    if os.name == 'nt':
        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()
            excel = None
            wb = None
            try:
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                excel.ScreenUpdating = False

                abs_path = os.path.abspath(pqp_path)
                wb = excel.Workbooks.Open(abs_path)

                for ws in wb.Worksheets:
                    sname = ws.Name.lower()
                    if '100%' in sname and 'dim' in sname:
                        last_row = ws.UsedRange.Rows.Count
                        for r in range(6, last_row + 1):
                            cell_b = ws.Cells(r, 2).Value
                            if cell_b is not None:
                                try:
                                    b_int = int(float(str(cell_b).strip()))
                                except Exception:
                                    b_int = str(cell_b).strip()

                                if b_int in target_3dec_balloons or str(b_int) in target_3dec_balloons:
                                    ws.Cells(r, 4).NumberFormat = "0.000"   # Nominal (Col D)
                                    ws.Cells(r, 5).NumberFormat = "0.000"   # Upper Tol (Col E)
                                    ws.Cells(r, 6).NumberFormat = "0.000"   # Lower Tol (Col F)
                                    ws.Cells(r, 7).NumberFormat = "0.000"   # Upper Spec (Col G)
                                    ws.Cells(r, 8).NumberFormat = "0.000"   # Lower Spec (Col H)
                                    ws.Cells(r, 9).NumberFormat = "0.0000"  # Measurement (Col I)
                                    repaired_count += 1
                                elif b_int in target_max_r_balloons or str(b_int) in target_max_r_balloons:
                                    ws.Cells(r, 5).NumberFormat = "0.00"
                                    ws.Cells(r, 7).NumberFormat = "0.00"
                                    ws.Cells(r, 8).NumberFormat = "0.00"
                                    repaired_count += 1

                wb.Save()
                com_success = True
                print(f"[Auto-Fix] Successfully repaired {repaired_count} cells via Excel Win32 COM!", flush=True)
            finally:
                if wb:
                    wb.Close(SaveChanges=False)
                if excel:
                    excel.Quit()
                pythoncom.CoUninitialize()
        except Exception as e:
            print(f"[Auto-Fix] Win32 COM error: {e}. Falling back to openpyxl...", flush=True)

    if not com_success:
        try:
            wb_fix = load_workbook_safe(pqp_path, data_only=False, keep_vba=False)
            repaired_count = 0
            for s in wb_fix.sheetnames:
                if '100%' in s.lower() and 'dim' in s.lower():
                    ws = wb_fix[s]
                    for r in range(6, ws.max_row + 1):
                        val_b = ws.cell(row=r, column=2).value
                        if val_b is not None:
                            try:
                                b_int = int(float(str(val_b).strip()))
                            except Exception:
                                b_int = str(val_b).strip()

                            if b_int in target_3dec_balloons or str(b_int) in target_3dec_balloons:
                                ws.cell(row=r, column=4).number_format = "0.000"
                                ws.cell(row=r, column=5).number_format = "0.000"
                                ws.cell(row=r, column=6).number_format = "0.000"
                                ws.cell(row=r, column=7).number_format = "0.000"
                                ws.cell(row=r, column=8).number_format = "0.000"
                                ws.cell(row=r, column=9).number_format = "0.0000"
                                repaired_count += 1
                            elif b_int in target_max_r_balloons or str(b_int) in target_max_r_balloons:
                                ws.cell(row=r, column=5).number_format = "0.00"
                                ws.cell(row=r, column=7).number_format = "0.00"
                                ws.cell(row=r, column=8).number_format = "0.00"
                                repaired_count += 1
            wb_fix.save(pqp_path)
            print(f"[Auto-Fix] Successfully repaired {repaired_count} cells via openpyxl!", flush=True)
        except Exception as e:
            print(f"[Auto-Fix] Error repairing workbook via openpyxl: {e}", file=sys.stderr)
            return False

    return True


def main():
    reconfigure_stdout()
    parser = argparse.ArgumentParser(
        description="Unified Drawing Specification, PQP Workbook, and MSA Quality Auditor (pqp-drawing-msa-auditor)"
    )
    subparsers = parser.add_subparsers(dest="command", help="Subcommand to run")
    
    # Subcommand: extract
    extract_parser = subparsers.add_parser("extract", help="Extract specifications and dimensions from drawing PDFs")
    extract_parser.add_argument("--dwgs", required=True, help="Path(s) to drawing PDF(s), comma-separated")
    extract_parser.add_argument("--output", default="spec.json", help="Path to write spec JSON")
    
    # Subcommand: audit
    audit_parser = subparsers.add_parser("audit", help="Audit PQP workbook compliance against specs / drawings")
    audit_parser.add_argument("--pqp", required=True, help="Path to PQP workbook (.xlsm or .xlsx)")
    audit_parser.add_argument("--spec", default="spec.json", help="Path to spec JSON file")
    audit_parser.add_argument("--dwgs", default=None, help="Optional path(s) to drawing PDF(s), comma-separated")
    audit_parser.add_argument("--output", default="audit_report.md", help="Path to write markdown compliance report")
    audit_parser.add_argument("--json-output", default=None, help="Optional path to write JSON audit report")
    audit_parser.add_argument("--fix-formats", action="store_true", help="Auto-fix outer PQP cell number format precision")
    
    # Subcommand: all
    all_parser = subparsers.add_parser("all", help="Extract specs and perform full 10-dimension PQP & Drawing audit in one step")
    all_parser.add_argument("--pqp", required=True, help="Path to PQP workbook (.xlsm or .xlsx)")
    all_parser.add_argument("--dwgs", default=None, help="Path(s) to drawing PDF(s), comma-separated")
    all_parser.add_argument("--spec", default=None, help="Path to pre-extracted spec JSON file (if dwgs not provided)")
    all_parser.add_argument("--output", default="audit_report.md", help="Path to write markdown compliance report")
    all_parser.add_argument("--json-output", default=None, help="Optional path to write JSON audit report")
    all_parser.add_argument("--fix-formats", action="store_true", help="Auto-fix outer PQP cell number format precision")

    # Subcommand: fix-formats
    fix_parser = subparsers.add_parser("fix-formats", help="Auto-fix cell number format precision in outer PQP dimensional results")
    fix_parser.add_argument("--pqp", required=True, help="Path to PQP workbook (.xlsm or .xlsx)")
    fix_parser.add_argument("--dwgs", default=None, help="Optional path(s) to drawing PDF(s), comma-separated")
    fix_parser.add_argument("--spec", default=None, help="Optional path to spec.json")

    # Subcommand: msa
    msa_parser = subparsers.add_parser("msa", help="Perform standalone Gage R&R and Attribute MSA numerical threshold audit")
    msa_parser.add_argument("--pqp", required=True, help="Path to PQP workbook (.xlsm or .xlsx)")
    msa_parser.add_argument("--output", default=None, help="Optional output report path")

    # Subcommand: checklist
    cl_parser = subparsers.add_parser("checklist", help="Perform standalone Index & PQS checklist alignment audit")
    cl_parser.add_argument("--pqp", required=True, help="Path to PQP workbook (.xlsm or .xlsx)")
    cl_parser.add_argument("--output", default=None, help="Optional output report path")

    # Subcommand: metadata
    meta_parser = subparsers.add_parser("metadata", help="Perform standalone metadata validation across all sheets")
    meta_parser.add_argument("--pqp", required=True, help="Path to PQP workbook (.xlsm or .xlsx)")
    meta_parser.add_argument("--spec", default=None, help="Optional path to spec.json")
    meta_parser.add_argument("--part-number", default=None, help="Expected Part Number")
    meta_parser.add_argument("--revision", default=None, help="Expected Revision")
    meta_parser.add_argument("--part-desc", default=None, help="Expected Part Description")
    meta_parser.add_argument("--drawing-number", default=None, help="Expected Drawing Number")
    
    # Subcommand: inspection
    insp_parser = subparsers.add_parser("inspection", help="Perform standalone Inspection Plan dimensional & SC balloon coverage audit")
    insp_parser.add_argument("--pqp", required=True, help="Path to PQP workbook (.xlsm or .xlsx)")
    insp_parser.add_argument("--output", default=None, help="Optional output report path")
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        sys.exit(1)
        
    if args.command == "extract":
        dwg_paths = parse_dwg_paths(args.dwgs)
        specs = {}
        for p in dwg_paths:
            if not os.path.exists(p):
                print(f"Error: Drawing file not found: {p}", file=sys.stderr)
                continue
            dp = DrawingParser(p)
            meta = dp.extract_metadata()
            torques = dp.extract_torques()
            sc_notes = dp.extract_sc_notes()
            sc_features = dp.extract_sc_features()
            dims_tols = dp.extract_dimensions_and_tolerances()
            
            dwg_key = meta["drawing_number"] or os.path.splitext(os.path.basename(p))[0]
            specs[dwg_key] = {
                "file": p,
                "metadata": meta,
                "torques": torques,
                "sc_notes": sc_notes,
                "sc_features": sc_features,
                "general_tolerance": dims_tols["general_tolerance"],
                "angle_tolerance": dims_tols["angle_tolerance"],
                "dimensions": dims_tols["dimensions"],
                "text": dims_tols["text"]
            }
            
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(specs, f, indent=2, ensure_ascii=False)
        print(f"Extracted specifications successfully written to {args.output}")
        return

    if args.command == "msa":
        if not os.path.exists(args.pqp):
            print(f"Error: PQP workbook file not found: {args.pqp}", file=sys.stderr)
            sys.exit(1)
        wb = load_workbook_safe(args.pqp, data_only=True, keep_vba=False)
        msa_anomalies = run_msa_and_capability_audit(wb, args.pqp)
        print("=== MSA & Process Capability Audit ===")
        if not msa_anomalies:
            print("PASS: All Gage R&R, Attribute MSA, and Capability metrics meet required thresholds.")
        else:
            print(f"FAIL: Found {len(msa_anomalies)} anomalies:")
            for a in msa_anomalies:
                print(f" - [{a['sheet']}] {a['metric']}: {a['reason']}")
        return

    if args.command == "checklist":
        if not os.path.exists(args.pqp):
            print(f"Error: PQP workbook file not found: {args.pqp}", file=sys.stderr)
            sys.exit(1)
        wb = load_workbook_safe(args.pqp, data_only=True, keep_vba=False)
        cl_mismatches = run_checklist_alignment_check(wb, args.pqp)
        print("=== Index & PQS Checklist Alignment Audit ===")
        if not cl_mismatches:
            print("PASS: Index and PQS checklists are perfectly aligned with actual worksheet content.")
        else:
            print(f"FAIL: Found {len(cl_mismatches)} checklist mismatches:")
            for m in cl_mismatches:
                print(f" - [{m['sheet']}] Row {m.get('row', 'N/A')} Col {m.get('col', 'N/A')}: {m['detail']}")
        return

    if args.command == "metadata":
        if not os.path.exists(args.pqp):
            print(f"Error: PQP workbook file not found: {args.pqp}", file=sys.stderr)
            sys.exit(1)
        wb = load_workbook_safe(args.pqp, data_only=True, keep_vba=False)
        try:
            wb_formulas = load_workbook_safe(args.pqp, data_only=False, keep_vba=False)
        except Exception:
            wb_formulas = None
        exp_dwg = args.drawing_number or args.part_number
        exp_rev = args.revision
        if args.spec and os.path.exists(args.spec):
            with open(args.spec, 'r', encoding='utf-8') as f:
                sdata = json.load(f)
                if isinstance(sdata, dict):
                    first_val = list(sdata.values())[0] if sdata else {}
                    if isinstance(first_val, dict) and "metadata" in first_val:
                        exp_dwg = first_val["metadata"].get("drawing_number", exp_dwg)
                        exp_rev = first_val["metadata"].get("revision", exp_rev)
                    elif "drawing_number" in sdata:
                        exp_dwg = sdata.get("drawing_number", exp_dwg)
                        exp_rev = sdata.get("revision", exp_rev)
        meta_mismatches = run_metadata_audit(wb, wb_formulas, exp_dwg, exp_rev)
        print("=== Metadata Audit ===")
        if not meta_mismatches:
            print("PASS: Metadata is consistent across all worksheets.")
        else:
            print(f"WARNING: Found {len(meta_mismatches)} metadata discrepancies:")
            for m in meta_mismatches:
                print(f" - [{m['sheet']}] {m['field']}: Expected '{m['expected']}', Found '{m['found']}'")
        return

    if args.command in ["inspection", "insp"]:
        if not os.path.exists(args.pqp):
            print(f"Error: PQP workbook file not found: {args.pqp}", file=sys.stderr)
            sys.exit(1)
        wb = load_workbook_safe(args.pqp, data_only=True, keep_vba=False)
        insp_res = run_inspection_plan_audit(wb)
        print("=== Inspection Plan Dimensional & SC Balloon Audit ===")
        print(f"Status: {insp_res['status']}")
        for f in insp_res["findings"]:
            print(f" - {f}")
        return

    if args.command == "fix-formats":
        if not os.path.exists(args.pqp):
            print(f"Error: PQP workbook file not found: {args.pqp}", file=sys.stderr)
            sys.exit(1)
        specs = {}
        if getattr(args, "dwgs", None):
            dwg_paths = parse_dwg_paths(args.dwgs)
            for p in dwg_paths:
                if not os.path.exists(p):
                    continue
                dp = DrawingParser(p)
                meta = dp.extract_metadata()
                dims_tols = dp.extract_dimensions_and_tolerances()
                dwg_key = meta["drawing_number"] or os.path.splitext(os.path.basename(p))[0]
                specs[dwg_key] = {
                    "dimensions": dims_tols["dimensions"],
                    "general_tolerance": dims_tols["general_tolerance"],
                    "metadata": meta
                }
        elif getattr(args, "spec", None) and os.path.exists(args.spec):
            with open(args.spec, 'r', encoding='utf-8') as f:
                specs = json.load(f)
        else:
            pqp_dir = os.path.dirname(os.path.abspath(args.pqp))
            auto_spec = os.path.join(pqp_dir, "spec.json")
            if os.path.exists(auto_spec):
                with open(auto_spec, 'r', encoding='utf-8') as f:
                    specs = json.load(f)

        dwg_spec_data = list(specs.values())[0] if specs else None
        wb = load_workbook_safe(args.pqp, data_only=True, keep_vba=False)
        cross_check_results = run_dimensional_spec_cross_check(wb, dwg_spec_data) if dwg_spec_data else None
        auto_fix_pqp_dimensional_formats(args.pqp, cross_check_results, dwg_spec_data)
        return

    # Subcommands 'all' and 'audit'
    if args.command in ["all", "audit"]:
        specs = {}
        if getattr(args, "dwgs", None):
            dwg_paths = parse_dwg_paths(args.dwgs)
            for p in dwg_paths:
                if not os.path.exists(p):
                    print(f"Error: Drawing file not found: {p}", file=sys.stderr)
                    continue
                dp = DrawingParser(p)
                meta = dp.extract_metadata()
                torques = dp.extract_torques()
                sc_notes = dp.extract_sc_notes()
                sc_features = dp.extract_sc_features()
                dims_tols = dp.extract_dimensions_and_tolerances()
                
                dwg_key = meta["drawing_number"] or os.path.splitext(os.path.basename(p))[0]
                specs[dwg_key] = {
                    "file": p,
                    "metadata": meta,
                    "torques": torques,
                    "sc_notes": sc_notes,
                    "sc_features": sc_features,
                    "general_tolerance": dims_tols["general_tolerance"],
                    "angle_tolerance": dims_tols["angle_tolerance"],
                    "dimensions": dims_tols["dimensions"],
                    "text": dims_tols["text"]
                }
        elif getattr(args, "spec", None) and os.path.exists(args.spec):
            with open(args.spec, 'r', encoding='utf-8') as f:
                specs = json.load(f)
        else:
            # Check if spec.json exists in target PQP directory
            pqp_dir = os.path.dirname(os.path.abspath(args.pqp))
            auto_spec = os.path.join(pqp_dir, "spec.json")
            if os.path.exists(auto_spec):
                print(f"Auto-detected spec.json in directory: {auto_spec}", flush=True)
                with open(auto_spec, 'r', encoding='utf-8') as f:
                    specs = json.load(f)
                
        # Load PQP Workbook
        if not os.path.exists(args.pqp):
            print(f"Error: PQP workbook file not found: {args.pqp}", file=sys.stderr)
            sys.exit(1)
            
        wb = load_workbook_safe(args.pqp, data_only=True, keep_vba=False)
        try:
            wb_formulas = load_workbook_safe(args.pqp, data_only=False, keep_vba=False)
        except Exception:
            wb_formulas = None
            
        # Match Drawing with PQP
        expected_dwg = None
        expected_rev = None
        dwg_spec_data = None
        single_dwg_specs = None
        
        if len(specs) == 1:
            k = list(specs.keys())[0]
            single_dwg_specs = specs[k]
            if isinstance(specs[k], dict) and "metadata" in specs[k]:
                expected_dwg = specs[k]["metadata"].get("drawing_number")
                expected_rev = specs[k]["metadata"].get("revision")
                dwg_spec_data = specs[k]
            elif isinstance(specs[k], dict):
                expected_dwg = specs[k].get("drawing_number")
                expected_rev = specs[k].get("revision")
                dwg_spec_data = {"metadata": specs[k], "torques": {}, "sc_notes": [], "text": ""}
            
        # Resolve target drawing
        pqp_dwg_cand = None
        pqp_rev_cand = None
        header_sheet = wb['TitlePg'] if 'TitlePg' in wb.sheetnames else (wb['PQS'] if 'PQS' in wb.sheetnames else None)
        if header_sheet:
            for r in range(1, min(30, header_sheet.max_row + 1)):
                for c in range(1, min(50, header_sheet.max_column + 1)):
                    val_s = str(header_sheet.cell(row=r, column=c).value or '').lower().strip()
                    if any(val_s == t or val_s.startswith(t + ":") for t in ["drawing number", "drawing #", "drawing no", "dwg no"]):
                        for co in range(1, 15):
                            pv = header_sheet.cell(row=r, column=c+co).value
                            if pv is not None and str(pv).strip():
                                pqp_dwg_cand = str(pv).strip()
                                break
                    elif any(val_s == t or val_s.startswith(t + ":") for t in ["revision", "rev", "rev."]):
                        for co in range(1, 15):
                            pv = header_sheet.cell(row=r, column=c+co).value
                            if pv is not None and str(pv).strip():
                                pqp_rev_cand = str(pv).strip()
                                break
            
        if pqp_dwg_cand and specs:
            norm_target = robust_norm(pqp_dwg_cand)
            dwg_found_match = False
            for k_spec, v_spec in specs.items():
                if isinstance(v_spec, dict):
                    v_meta = v_spec.get("metadata", v_spec)
                    if robust_norm(k_spec) == norm_target or (v_meta and robust_norm(v_meta.get("drawing_number")) == norm_target):
                        expected_dwg = v_meta.get("drawing_number")
                        expected_rev = v_meta.get("revision")
                        dwg_spec_data = v_spec if "metadata" in v_spec else {"metadata": v_spec, "torques": {}, "sc_notes": [], "text": ""}
                        dwg_found_match = True
                        break
                    
            if not dwg_found_match and specs:
                first_key = list(specs.keys())[0]
                dwg_spec_data = specs[first_key]
                if "metadata" not in dwg_spec_data and isinstance(dwg_spec_data, dict):
                    dwg_spec_data = {"metadata": dwg_spec_data, "torques": {}, "sc_notes": [], "text": ""}
                
        if not dwg_spec_data and single_dwg_specs:
            dwg_spec_data = single_dwg_specs
            
        # Prioritize authoritative drawing specification from drawing PDF or spec JSON if provided
        dwg_authoritative_no = None
        dwg_authoritative_rev = None
        dwg_authoritative_title = None
        if dwg_spec_data and isinstance(dwg_spec_data, dict):
            m = dwg_spec_data.get("metadata", {})
            if isinstance(m, dict):
                dwg_authoritative_no = m.get("drawing_number")
                dwg_authoritative_rev = m.get("revision")
                dwg_authoritative_title = m.get("title")

        if dwg_authoritative_no and str(dwg_authoritative_no).strip().upper() not in ["ITER", "N/A"]:
            expected_dwg = dwg_authoritative_no
        elif pqp_dwg_cand:
            expected_dwg = pqp_dwg_cand
        elif expected_dwg is None:
            expected_dwg = "N/A"
            
        if dwg_authoritative_rev and str(dwg_authoritative_rev).strip().upper() not in ["N/A"]:
            expected_rev = dwg_authoritative_rev
        elif pqp_rev_cand:
            expected_rev = pqp_rev_cand
        elif expected_rev is None:
            expected_rev = "N/A"
            
        expected_desc = dwg_authoritative_title
        expected_pn = expected_dwg
        
        try:
            expected_dwg = int(float(str(expected_dwg).strip()))
        except ValueError:
            pass
        try:
            expected_rev = int(float(str(expected_rev).strip()))
        except ValueError:
            pass
        
        # Run Audits
        print("[Step 1/10] Running metadata audit...", flush=True)
        meta_mismatches = run_metadata_audit(wb, wb_formulas, expected_dwg, expected_rev, expected_desc=expected_desc, expected_pn=expected_pn)
        print("[Step 2/10] Running Index & PQS checklist alignment audit...", flush=True)
        checklist_mismatches = run_checklist_alignment_check(wb, args.pqp)
        print("[Step 3/10] Running control plan torque audit...", flush=True)
        torque_mismatches = run_control_plan_audit(wb)
        print("[Step 4/10] Running control plan CTQ/SC audit...", flush=True)
        ctq_sc_audit_res = run_control_plan_ctq_sc_audit(wb, dwg_spec_data)
        print("[Step 4.1/10] Running inspection plan dimensional & SC balloon audit...", flush=True)
        insp_audit_res = run_inspection_plan_audit(
            wb,
            dwg_spec_data,
            dim_sc_items=ctq_sc_audit_res.get("sc_ctq_dict"),
            all_dim_balloons=ctq_sc_audit_res.get("all_dim_balloons")
        )
        print("[Step 5/10] Running MSA and process capability audit...", flush=True)
        msa_anomalies = run_msa_and_capability_audit(wb, args.pqp, dwg_spec_data)
        print("[Step 6/10] Running 100% dimensional inspection audit...", flush=True)
        dim_warnings, dim_sheets = run_dimensional_audit(wb)
        print("[Step 7/10] Running dimensional spec cross check...", flush=True)
        cross_check_results = run_dimensional_spec_cross_check(wb, dwg_spec_data)
        if getattr(args, "fix_formats", False):
            print("[Auto-Fix] Running cell number format repair for outer PQP dimensional results...", flush=True)
            auto_fix_pqp_dimensional_formats(args.pqp, cross_check_results, dwg_spec_data)
            try:
                wb = load_workbook_safe(args.pqp, data_only=True, keep_vba=False)
                cross_check_results = run_dimensional_spec_cross_check(wb, dwg_spec_data)
                print("[Auto-Fix] Reloaded PQP workbook and updated cross-check findings.", flush=True)
            except Exception as e:
                print(f"[Auto-Fix] Note: Could not reload workbook: {e}", flush=True)
        print("[Step 8/10] Running 3F/FFF photo audit...", flush=True)
        dwg_text_for_fff = dwg_spec_data.get("text", "") if dwg_spec_data else ""
        fff_mismatches = run_fff_photo_audit(args.pqp, dwg_text_for_fff, wb=wb, dwg_spec_data=dwg_spec_data)
        print("[Step 9/10] Running embedded FAI & test certificates audit...", flush=True)
        raw_fai_discrepancies = run_embedded_fai_discrepancy_audit(args.pqp, wb)
        embedded_certs = run_embedded_attachments_audit(args.pqp, wb=wb, dwg_spec_data=dwg_spec_data)
        print("[Step 10/10] Running unrelated embedded files audit...", flush=True)
        unrelated_findings, own_pns_list = run_unrelated_embeddings_audit(args.pqp, wb)
        print("Audits complete. Generating comprehensive report...", flush=True)
        
        # Build Report
        report_md = []
        report_md.append("# Drawing Specification, PQP Workbook & MSA Quality Compliance Report\n")
        report_md.append(f"**Target Workbook**: `{os.path.basename(args.pqp)}`\n")
        
        if 'TitlePg' in wb.sheetnames:
            selected_template = wb['TitlePg'].cell(row=8, column=4).value
        elif 'PQS' in wb.sheetnames:
            selected_template = wb['PQS'].cell(row=1, column=16).value
        else:
            selected_template = "Unknown Template"
        report_md.append(f"* **Active Template**: `{selected_template}`")
        visible_sheets = [s.title for s in wb.worksheets if s.sheet_state == 'visible']
        report_md.append(f"* **Active Worksheets**: {len(visible_sheets)} visible sheets\n")
        report_md.append("---")
        
        # Section 1: Metadata Check
        report_md.append("\n## 1. PQP Metadata Integrity & Family Part Check")
        if not meta_mismatches:
            report_md.append("✅ **PASS**: All visible sheets have consistent, correct metadata headers matching drawing specifications.")
        else:
            report_md.append(f"⚠️ **WARNING**: Found {len(meta_mismatches)} inconsistent metadata values in visible sheets:")
            for m in meta_mismatches:
                report_md.append(f"- Worksheet `[{m['sheet']}]`: Field `{m['field']}` mismatch. Expected: `{m['expected']}`, Found: `{m['found']}`")

        # Section 2: Index & PQS Checklist Alignment Check
        report_md.append("\n## 2. Index & PQS Checklist Alignment Audit")
        if not checklist_mismatches:
            report_md.append("✅ **PASS**: All checklist markings in `Index` status and `PQS` cover page match actual worksheet content and data activity.")
        else:
            report_md.append(f"❌ **FAIL / WARNING**: Found {len(checklist_mismatches)} checklist alignment discrepancies:")
            for cl in checklist_mismatches:
                pos_info = f" (Row {cl['row']}, Col {cl['col']})" if cl.get('row') else ""
                report_md.append(f"- **Worksheet `[{cl['sheet']}]`**{pos_info}: {cl['detail']}")
                
        # Section 3: Torque Specifications Audit
        report_md.append("\n## 3. Control Plan Torque Specifications Audit")
        if not torque_mismatches:
            report_md.append("✅ **PASS**: All assembly torques in the Control Plan are compliant with drawing flag notes.")
        else:
            report_md.append(f"❌ **FAIL**: Discovered {len(torque_mismatches)} mismatches in Control Plan assembly torques:")
            for t in torque_mismatches:
                report_md.append(f"- **Step {t['step']} (Row {t['row']})**: {t['reason']}. Found: `{t['found_torque']} N-m` in description.")

        # Section 3.1: Control Plan CTQ / SC Coverage & Traceability Audit
        report_md.append("\n## 3.1 Control Plan CTQ / SC Coverage & Traceability Audit")
        sc_items = ctq_sc_audit_res.get("sc_ctq_items", [])
        dwg_sc_notes = ctq_sc_audit_res.get("dwg_sc_notes", [])
        dwg_sc_features = ctq_sc_audit_res.get("dwg_sc_features", [])
        
        if dwg_sc_notes:
            report_md.append(f"ℹ️ **圖面標註之 SC 重點附註 (Drawing SC Flag Notes: {len(dwg_sc_notes)} 項)**:")
            for sc in dwg_sc_notes:
                report_md.append(f"- **Note {sc['note_num']}**: `◇ {sc['content']}`")
                
        if dwg_sc_features:
            report_md.append(f"\nℹ️ **工程圖面菱形 ◇ / SC 標記特徵清單 (Drawing SC Diamond Features: {len(dwg_sc_features)} 項)**:")
            for sc in dwg_sc_features:
                b_str = f" -> 對應尺寸表 Balloon #{sc['matched_balloon']}" if sc.get("matched_balloon") else ""
                report_md.append(f"- **特徵代號 `{sc['letter']}`**: Nominal `{sc['nominal']}` ({sc.get('raw_text', '')}){b_str}")
                
        if sc_items:
            report_md.append(f"\nℹ️ **尺寸表宣告之管制特徵清單 (Declared SC/CTQ Items: {len(sc_items)} 項)**:")
            for sc in sc_items:
                report_md.append(f"- **Balloon #{sc['balloon']}** [{sc['type']}]: Nominal: `{sc['nominal']}` (Tol: `+{sc['tol_up']}/-{sc['tol_lo']} {sc['unit']}`), 預定量具: `{sc['tool'] or '未指定'}` (來源表: `{sc['sheet']}`)")
        
        if not ctq_sc_audit_res.get("findings"):
            report_md.append("✅ **PASS**: All CTQ / SC characteristics are fully traced and covered in the Control Plan, and sampling inspection reaction plans satisfy statistical process control (containment & sorting suspect units produced since last inspection).")
        else:
            for f in ctq_sc_audit_res["findings"]:
                report_md.append(f"- {f}")
                
        # Section 3.2: Inspection Plan Dimensional & SC Balloon Coverage Audit
        report_md.append("\n## 3.2 Inspection Plan Dimensional & SC Balloon Coverage Audit")
        if insp_audit_res["status"] == "PASS":
            report_md.append("✅ **PASS**: Inspection Plan properly references and traces all critical SC dimensions and balloons.")
        elif insp_audit_res["status"] == "WARNING":
            report_md.append("⚠️ **WARNING**: Inspection Plan has traceability observations or generic descriptions:")
        else:
            report_md.append(f"❌ **FAIL**: Found {len(insp_audit_res.get('sc_balloons_missing_in_insp', []))} missing SC balloon(s) or critical omissions in Inspection Plan:")
            
        for f in insp_audit_res.get("findings", []):
            report_md.append(f"- {f}")
            
        if insp_audit_res.get("sc_balloons_declared"):
            sc_strs = [f"Balloon #{b}" for b in insp_audit_res["sc_balloons_declared"]]
            report_md.append(f"\nℹ️ **圖面/尺寸表已宣告之 SC 管制球標 ({len(sc_strs)} 項)**: `{', '.join(sc_strs)}`")
        if insp_audit_res.get("sc_balloons_in_insp"):
            sc_in_strs = [f"Balloon #{b}" for b in insp_audit_res["sc_balloons_in_insp"]]
            report_md.append(f"ℹ️ **Inspection Plan 已標示之 SC 球標 ({len(sc_in_strs)} 項)**: `{', '.join(sc_in_strs)}`")
        if insp_audit_res.get("sc_balloons_missing_in_insp"):
            sc_miss_strs = [f"Balloon #{b}" for b in insp_audit_res["sc_balloons_missing_in_insp"]]
            report_md.append(f"❌ **Inspection Plan 缺漏之 SC 球標 ({len(sc_miss_strs)} 項)**: `{', '.join(sc_miss_strs)}`")

        # Section 4: MSA & Process Capability Audit
        report_md.append("\n## 4. MSA (Gage R&R & Attribute) & Process Capability Study Audit")
        if not msa_anomalies:
            report_md.append("✅ **PASS**: All Gage R&R, Attribute MSA, and capability metrics satisfy quality thresholds (% Contribution < 9%, % Study Var < 20%, % Tolerance < 20%, ndc > 4, Attribute Effective Scores = 100%).")
        else:
            report_md.append(f"❌ **FAIL**: Found {len(msa_anomalies)} anomalies or threshold violations:")
            for a in msa_anomalies:
                report_md.append(f"- **Worksheet `[{a['sheet']}]`** ({a['metric']}): {a['reason']}")
                
        # Section 5: 100% Dimensional Inspection Audit
        report_md.append("\n## 5. 100% Dimensional Inspection Audit")
        if not dim_sheets:
            report_md.append("ℹ️ **N/A**: No active dimensional inspection sheets found or required for this template.")
        elif not dim_warnings:
            report_md.append("✅ **PASS**: All dimensional measurements and Sample # contiguous sequences are within tolerance limits with no drift alerts.")
        else:
            report_md.append("⚠️ **WARNING**: Flagged drift alerts or border-line measurements (disposition is not 'Acceptable'):")
            by_sheet = {}
            for w in dim_warnings:
                by_sheet.setdefault(w["sheet"], []).append(w)
            for sname, items in by_sheet.items():
                report_md.append(f"\n### Worksheet `[{sname}]`:")
                for item in items:
                    report_md.append(f"- **Row {item['row']} (Balloon {item['balloon']})**: Nominal: `{item['nominal']}`, Measured: `{item['measured']}`. Disposition: `{item['disposition']}`")
                    
        # Section 6: Drawing vs PQP Dimensional Spec Cross-Reference Audit
        report_md.append("\n## 6. Drawing vs PQP Dimensional Spec Cross-Reference Audit")
        failures = [r for r in cross_check_results if r.get("status") == "FAIL"]
        warnings_cc = [r for r in cross_check_results if r.get("status") == "WARNING"]
        passes = [r for r in cross_check_results if r.get("status") == "PASS"]
        
        if not cross_check_results or (len(cross_check_results) == 1 and cross_check_results[0]["status"] == "INFO"):
            report_md.append("ℹ️ **N/A**: No drawing spec comparison data available.")
        else:
            if failures:
                report_md.append(f"❌ **FAIL**: Discovered {len(failures)} dimensional discrepancies between PDF drawing and Excel PQP:")
                for f in failures:
                    report_md.append(f"- **Worksheet `[{f['sheet']}]` Balloon {f['balloon']}**: {f['message']}")
            else:
                report_md.append("✅ **PASS**: All Excel nominals and tolerances match the PDF drawing specifications.")
                
            if warnings_cc:
                report_md.append(f"\n⚠️ **WARNING ({len(warnings_cc)} qualitative or extra match issues)**:")
                for w in warnings_cc:
                    report_md.append(f"- **Worksheet `[{w['sheet']}]` Balloon {w['balloon']}**: {w['message']}")
            
            report_md.append(f"\n### Detail Comparison ({len(passes)} items matched successfully):")
            for p in passes:
                report_md.append(f"- **Balloon {p['balloon']}**: {p['message']}")
                
        # Section 7: Form, Fit, Function (3F) Photo & Content Compliance Audit
        report_md.append("\n## 7. Form, Fit, Function (3F) Photo & Content Compliance Audit")
        failures = [m for m in fff_mismatches if m.get("status") == "FAIL"]
        warnings = [m for m in fff_mismatches if m.get("status") in ["WARNING", "ERROR"]]
        infos = [m for m in fff_mismatches if m.get("status") == "INFO"]
        passes = [m for m in fff_mismatches if m.get("status") == "PASS"]
        
        if failures:
            report_md.append(f"❌ **FAIL**: Found {len(failures)} mismatching parameter(s) / defect(s) in 3F (FFF) verification table & photos:")
            for f in failures:
                if "message" in f:
                    report_md.append(f"- {f['message']}")
                else:
                    report_md.append(f"- **Photo {f.get('photo_num', 'N/A')}**: Field `{f.get('field', 'N/A')}` mismatch. Expected (Drawing): `{f.get('expected')}`, Found (Photo): `{f.get('found')}`")
        if warnings:
            report_md.append(f"⚠️ **WARNING**: Issues encountered during 3F (FFF) audit:")
            for w in warnings:
                if "message" in w:
                    report_md.append(f"- {w['message']}")
                else:
                    report_md.append(f"- **Photo {w.get('photo_num', 'Unknown')}**: Field `{w.get('field', 'Unknown')}` warning. Expected: `{w.get('expected')}`, Found: `{w.get('found')}`")
        if infos:
            report_md.append("\n### 檢驗說明與規範建議 (Advisories):")
            for inf in infos:
                report_md.append(f"- {inf['message']}")
        if passes:
            report_md.append("\n### 驗證通過項目 (Verified Items):")
            for p in passes:
                report_md.append(f"- {p['message']}")
        if not failures and not warnings and not infos and not passes:
            report_md.append("✅ **PASS**: All 3F (FFF) verification reports and photos are compliant with drawing specifications.")

        # Section 8: Embedded Raw FAI & Test Certificates Audit
        report_md.append("\n## 8. Embedded Raw FAI & Test Certificates Audit")
        fai_errs = [d for d in raw_fai_discrepancies if d.get('status') == 'FAIL']
        fai_warns = [d for d in raw_fai_discrepancies if d.get('status') == 'WARNING']
        fai_infos = [d for d in raw_fai_discrepancies if d.get('status') == 'INFO']
        
        if fai_errs:
            report_md.append(f"❌ **CRITICAL DISCREPANCY DETECTED**: Found {len(fai_errs)} discrepancies between active PQP sheet and embedded raw FAI sub-workbook:")
            for d in fai_errs:
                report_md.append(f"- **[{d['status']}]**: {d['message']}")
        elif fai_warns:
            report_md.append(f"⚠️ **WARNING**: Found {len(fai_warns)} warnings between active PQP sheet and embedded raw FAI sub-workbook:")
            for d in fai_warns:
                report_md.append(f"- **[{d['status']}]**: {d['message']}")
        else:
            report_md.append("✅ **PASS**: Embedded raw FAI records are consistent with active PQP sheets.")
            
        if fai_infos and not fai_errs and not fai_warns:
            for d in fai_infos:
                report_md.append(f"- **[{d['status']}]**: {d['message']}")
            
        if embedded_certs:
            report_md.append(f"\n### 附加檢驗與認證報告清單 (Attached Test Reports: {len(embedded_certs)} files):")
            for c in embedded_certs:
                desc_info = ""
                if c.get("desc_status") == "mismatch":
                    desc_info = f"\n  - ⚠️ **[品名不相符 WARNING]**: {c['desc_msg']}"
                elif c.get("desc_status") == "match":
                    desc_info = f"\n  - ✅ **[品名相符]**: {c['desc_msg']}"
                report_md.append(f"- **分頁 `{c.get('sheet_str', '未知分頁')}`** | **`{c['file']}`** ({c['pages']} 頁): {c['type']} - {c['summary']}{desc_info}")
        else:
            report_md.append("\nℹ️ **No embedded OLE test reports found.**")

        # Section 9: Unrelated Embedded Files Audit
        report_md.append("\n## 9. Unrelated Embedded Files Audit (Copy-Paste Residual Check)")
        if own_pns_list:
            report_md.append(f"ℹ️ **本 PQP 已知料號 (Own PNs)**: `{', '.join(own_pns_list)}`")
        unrelated_items = [f for f in unrelated_findings if f['status'] == 'unrelated']
        ok_items = [f for f in unrelated_findings if f['status'] == 'ok']
        no_pn_items = [f for f in unrelated_findings if f['status'] == 'no_pn']

        if not unrelated_findings:
            report_md.append("ℹ️ **N/A**: 本工作簿內未偵測到任何嵌入物件。")
        elif not unrelated_items:
            report_md.append("✅ **PASS**: 所有嵌入物件的料號均與本 PQP 相符，未發現誤植的其他產品文件。")
        else:
            report_md.append(f"❌ **FAIL**: 發現 {len(unrelated_items)} 個嵌入物件的料號與本 PQP 不符，可能為誤植的其他產品文件：")
            for item in unrelated_items:
                report_md.append(
                    f"- **分頁 `{item.get('sheet_str', '未知分頁')}`** | **`{item['file']}`** ({item['size_kb']:.0f} KB, [{item['type']}]): "
                    f"偵測料號 `{', '.join(item['pns_found'])}` — {item['message']}"
                )
                if item['preview']:
                    report_md.append(f"  - 內容預覽: `{item['preview'][:200]}`")

        if ok_items:
            report_md.append(f"\n### 相符嵌入物件 ({len(ok_items)} 個):")
            for item in ok_items:
                report_md.append(f"- **分頁 `{item.get('sheet_str', '未知分頁')}`** | **`{item['file']}`** ({item['size_kb']:.0f} KB): {item['message']}")

        if no_pn_items:
            report_md.append(f"\n### 無法偵測料號之嵌入物件 ({len(no_pn_items)} 個，需人工確認):")
            for item in no_pn_items:
                report_md.append(
                    f"- **分頁 `{item.get('sheet_str', '未知分頁')}`** | **`{item['file']}`** ({item['size_kb']:.0f} KB, [{item['type']}]): {item['message']}"
                )

        # Write markdown report
        try:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write("\n".join(report_md))
            print(f"Compliance report successfully written to {args.output}")
        except Exception as e:
            print(f"Error writing markdown report: {e}", file=sys.stderr)
            sys.exit(1)
            
        # Write JSON report if requested
        if getattr(args, "json_output", None):
            json_data = {
                "pqp_file": args.pqp,
                "metadata_mismatches": meta_mismatches,
                "checklist_mismatches": checklist_mismatches,
                "torque_mismatches": torque_mismatches,
                "ctq_sc_audit": ctq_sc_audit_res,
                "inspection_plan_audit": insp_audit_res,
                "msa_anomalies": msa_anomalies,
                "dimensional_warnings": dim_warnings,
                "cross_check_results": cross_check_results,
                "fff_mismatches": fff_mismatches,
                "raw_fai_discrepancies": raw_fai_discrepancies,
                "embedded_certs": embedded_certs,
                "unrelated_findings": unrelated_findings
            }
            try:
                with open(args.json_output, 'w', encoding='utf-8') as jf:
                    json.dump(json_data, jf, indent=2, ensure_ascii=False)
                print(f"JSON compliance report successfully written to {args.json_output}")
            except Exception as e:
                print(f"Error writing JSON report: {e}", file=sys.stderr)

if __name__ == "__main__":
    main()
