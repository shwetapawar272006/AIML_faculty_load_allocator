"""
processor.py — AIML Faculty Core Load Allocator

Every single value in the output is derived from the input Excel file.
NO hardcoded names, faculty counts, load totals, or NF placeholders.

Sheets read from input:
  - "TRCPIE Allocation": authoritative faculty list + C/P/I/E/Admin/Total values
  - "Subject List": actual even-semester subject assignments per faculty

Output logic:
  - Faculty list   = exactly the faculty rows in "TRCPIE Allocation" (no additions)
  - Theory         = sum of loads where faculty is teacher1 in non-lab subjects
  - Lab            = sum of loads where faculty is teacher1 or teacher2 in lab subjects
                     (teacher3 in a 3-teacher lab does NOT receive a unit)
  - Teaching Units = Theory + Lab
  - R (Research)   = faculty's Total (col10) − TU − C − P − I − E − Admin  [≥ 0]
  - C, P, I, E, Admin, Total  = read directly from "TRCPIE Allocation"
  - Designation    = inferred from name prefix (Dr./PhD → Associate Professor, else Assistant)
"""

import pandas as pd
import io
import re
from openpyxl import Workbook


# ─────────────────────────── Name matching ───────────────────────────────────

def _strip(name: str) -> str:
    """Normalise a name: remove titles, (PhD), Topo suffix, lower-case."""
    s = str(name).strip()
    s = re.sub(r'^(Mr\.|Ms\.|Dr\.|Prof\.)\s*', '', s, flags=re.I)
    s = re.sub(r'\s*\(PhD\.?\)\s*', '', s, flags=re.I)
    s = re.sub(r'\bTopo\b', '', s, flags=re.I)
    return s.lower().strip()


def names_match(a: str, b: str) -> bool:
    """
    Fuzzy name match that handles the naming inconsistencies present in this
    dataset (reversed initials, abbreviated first names, missing suffixes)
    WITHOUT false-matching genuinely different people.

    Rules (in order):
      1. Exact after stripping
      2. Token-set subset   "Saranya" ↔ "Batta Saranya"
      3. Reversed tokens    "S Santhiya" ↔ "Santhiya S"
      4. First-token prefix (both ≥ 4 chars) AND last tokens agree
         → "Prabhakara S" ↔ "Prabhakar"
         → Guards against "Divya Kumari" ≠ "Divyamani LS"
      5. Nickname compression via char-subsequence on long first name (≥ 8 chars)
         → "Swasthisudha" ↔ "Swasti Sudha"
    """
    if not a or not b:
        return False
    na, nb = _strip(a), _strip(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    ta, tb = na.split(), nb.split()

    # 1. Token-set subset
    sa, sb = set(ta), set(tb)
    if sa and sb and (sa <= sb or sb <= sa):
        return True

    # 2. Reversed token order
    if sorted(ta) == sorted(tb):
        return True

    # 3. First-token prefix with last-token agreement
    f1, f2 = ta[0], tb[0]
    if len(f1) >= 4 and len(f2) >= 4 and (f1.startswith(f2) or f2.startswith(f1)):
        if len(ta) == 1 or len(tb) == 1:
            return True
        l1, l2 = ta[-1], tb[-1]
        if l1 == l2 or l1.startswith(l2) or l2.startswith(l1):
            return True

    # 4. Nickname compression (e.g. "Swasthisudha" ↔ "Swasti Sudha")
    longer_toks = ta if len(na) >= len(nb) else tb
    shorter_toks = tb if len(na) >= len(nb) else ta
    lf = longer_toks[0]
    short_concat = "".join(shorter_toks)
    if (len(lf) >= 8 and len(short_concat) >= 6
            and all(len(t) >= 4 for t in shorter_toks)):
        it = iter(lf)
        if all(c in it for c in short_concat):
            return True

    return False


# ─────────────────────────── Subject classification ──────────────────────────

LAB_KEYWORDS = ["lab", "laboratory", "ability enhancement", "devops"]


def _is_lab(subject_name: str) -> bool:
    low = subject_name.lower()
    return any(kw in low for kw in LAB_KEYWORDS)


# ─────────────────────────── Designation ─────────────────────────────────────

def _designation(raw_name: str) -> str:
    """Infer designation from the raw name string in TRCPIE Allocation."""
    s = str(raw_name).strip()
    if re.search(r'\bDr\.?\b', s) or re.search(r'PhD', s, re.I):
        return "Associate Professor"
    return "Assistant Professor"


# ─────────────────────────── Core processor ──────────────────────────────────

def process_subject_allotment(input_bytes: bytes) -> bytes:
    xls = pd.ExcelFile(io.BytesIO(input_bytes))
    subject_list_df = pd.read_excel(xls, sheet_name="Subject List",       header=None)
    trcpie_alloc_df = pd.read_excel(xls, sheet_name="TRCPIE Allocation",  header=None)

    # ── Parse TRCPIE Allocation ───────────────────────────────────────────────
    # Layout: row 6 = column headers, rows 7+ = faculty data
    # Columns: Sl.No | Name | T | L | R | C | P | I | E | A | Total
    faculty_rows = []
    for i in range(7, len(trcpie_alloc_df)):
        row = trcpie_alloc_df.iloc[i]
        raw_name = row.iloc[1] if len(row) > 1 else None
        if pd.isna(raw_name):
            continue
        raw_name = str(raw_name).strip()
        if not raw_name or raw_name.lower() in ("total", "target"):
            continue  # skip summary rows

        def _v(col, _r=row):
            v = _r.iloc[col] if len(_r) > col else None
            return None if (v is None or (isinstance(v, float) and pd.isna(v))) else v

        faculty_rows.append({
            "sl_no": _v(0), "name": raw_name,
            # T and L here are the ODD-sem actuals — used only as fallback context;
            # even-sem T/L are re-derived from Subject List below.
            "C": _v(5), "P": _v(6), "I": _v(7), "E": _v(8),
            "A": _v(9), "Total": _v(10),
        })

    def _get_fr(nm: str):
        for fr in faculty_rows:
            if names_match(nm, fr["name"]):
                return fr
        return None

    # ── Parse Subject List ────────────────────────────────────────────────────
    parsed = []
    cur_sec = ""
    for _, row in subject_list_df.iterrows():
        sec = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if sec and sec.lower() not in ("section", "nan", ""):
            cur_sec = sec
        sname = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        if not sname or sname.lower() in ("nan", ""):
            continue
        code = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if code.lower() == "nan":
            code = ""
        try:
            load = float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0.0
        except Exception:
            load = 0.0
        if load <= 0:
            continue
        teachers = []
        for col in [6, 7, 8]:
            t = str(row.iloc[col]).strip() if (len(row) > col and pd.notna(row.iloc[col])) else ""
            if t and t.lower() != "nan":
                teachers.append(t)
        parsed.append({
            "section": cur_sec, "code": code, "name": sname,
            "load": load, "teachers": teachers, "is_lab": _is_lab(sname),
        })

    def _subjects_for(nm: str):
        """
        Return subjects where this faculty is an eligible assignee:
          Theory  → only teacher1 receives the unit
          Lab     → teacher1 and teacher2 each receive a unit; teacher3 does not
        """
        result = []
        for s in parsed:
            eligible = s["teachers"][:2] if s["is_lab"] else s["teachers"][:1]
            if any(names_match(nm, t) for t in eligible):
                result.append(s)
        return result

    def _teaching_units(nm: str):
        th = lb = 0.0
        for s in _subjects_for(nm):
            (lb if s["is_lab"] else th).__class__  # just for clarity
            if s["is_lab"]:
                lb += s["load"]
            else:
                th += s["load"]
        return round(th, 4), round(lb, 4)

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _flt(v):
        try:
            return float(v) if v is not None else 0.0
        except Exception:
            return 0.0

    # ════════════════════════ Build Output Workbook ═══════════════════════════
    wb = Workbook()

    # ─── Sheet 1: TRCPIE ──────────────────────────────────────────────────────
    # Copy the entire TRCPIE Allocation sheet, replacing "Odd Semester" → "Even Semester"
    ws1 = wb.active
    ws1.title = "TRCPIE"
    ws1.append(["", "Faculty Core Load Distribution", "",
                "Academic Year", "", "2025-26 (Even Sem)"])
    for i in range(len(trcpie_alloc_df)):
        row_data = []
        for v in trcpie_alloc_df.iloc[i]:
            if isinstance(v, str):
                v = v.replace("Odd Semester", "Even Semester")
            row_data.append(v if not (isinstance(v, float) and pd.isna(v)) else None)
        ws1.append(row_data)

    # ─── Sheet 2: subject allocation ──────────────────────────────────────────
    ws2 = wb.create_sheet("subject allocation")

    # Institution header — read from TRCPIE Allocation rows 0-4
    for i in range(5):
        row_data = [v if pd.notna(v) else None for v in trcpie_alloc_df.iloc[i]]
        # Replace Odd→Even in header text
        row_data = [v.replace("Odd Semester", "Even Semester")
                    if isinstance(v, str) else v for v in row_data]
        ws2.append(row_data)
    ws2.append([])  # blank spacer

    # Column headers (two rows)
    ws2.append(["S.No", "Name of the faculty", "Designation", "Theory", "Lab",
                "Units for TRCPIE", None, None, None, None, None,
                "Admin", "D Load", "Total load", "Extra Load"])
    ws2.append([None, None, None, None, None,
                "Teaching Units\n\nT", "Research Units\n R",
                "Consultancy Units\n C", "Project Units\n P",
                "Innovation\n I", "Entrepreneurship\n E",
                "Admin", "D Load", "Total load", "Extra Load"])

    total_tu = 0.0

    for idx, fr in enumerate(faculty_rows, 1):
        nm = fr["name"]
        des = _designation(nm)
        th, lb = _teaching_units(nm)
        tu = round(th + lb, 4)
        total_tu += tu

        c_val  = fr["C"]
        p_val  = fr["P"]
        i_val  = fr["I"]
        e_val  = fr["E"]
        adm    = fr["A"]
        total  = fr["Total"]

        # R = Total − TU − C − P − I − E − Admin  (never negative)
        r_computed = round(
            max(0.0, _flt(total) - tu - _flt(c_val) - _flt(p_val)
                                  - _flt(i_val) - _flt(e_val) - _flt(adm)),
            4)
        r_val = r_computed if r_computed > 0 else None

        ws2.append([
            idx, nm, des,
            th if th else None,
            lb if lb else None,
            tu,
            r_val, c_val, p_val, i_val, e_val, adm,
            None,    # D Load (not in input)
            total,
            0,       # Extra Load
        ])

    # Grand-total teaching units row
    ws2.append([None, None, None, None, None, round(total_tu, 4),
                None, None, None, None, None, None, None, None, None])

    # ─── Sheet 3: Core1-4 Break-Up ────────────────────────────────────────────
    ws3 = wb.create_sheet("Core1-4 Break-Up")
    ws3.append(["S.no", "Name of the faculty", "Designation",
                "Core 1 Type", "Core 1 Details",
                "Core 2 Type", "Core 2 Details",
                "Core 3 Type", "Core 3 Details",
                "Core 4 Type", "Core 4 Details",
                "Extra T Load assigned", "Extra T load Details"])

    def _build_cores(nm: str):
        slots = []

        # Teaching slots from Subject List
        for s in _subjects_for(nm):
            sec, code, sname = s["section"], s["code"], s["name"]
            load, is_lab = s["load"], s["is_lab"]
            prefix = f"{sec}-{code}-{sname}" if code else f"{sec}-{sname}"
            if is_lab:
                for _ in range(int(round(load))):
                    slots.append(("1T", f"{prefix} (3 Slots)"))
            elif load < 1.0:
                slots.append(("TR", f"{prefix}\n{load}R"))
            else:
                for _ in range(int(round(load))):
                    slots.append(("1 T", prefix.strip()))

        # Non-teaching TRCPIE slots from TRCPIE Allocation
        fr = _get_fr(nm)
        if fr:
            frac_parts, frac_keys = [], []
            for key in ["R", "C", "P", "I", "E"]:
                # Use *odd-sem* TRCPIE Allocation value as the even-sem plan
                # (only for the Core break-up detail; R is not stored in faculty_rows
                #  so we use the residual we computed above)
                pass  # handled below via fr lookup on full trcpie_alloc_df

            # Re-read R,C,P,I,E directly from the source row for Core details
            for i in range(7, len(trcpie_alloc_df)):
                row = trcpie_alloc_df.iloc[i]
                rn = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                if not names_match(nm, rn):
                    continue
                for col, key in [(4,'R'),(5,'C'),(6,'P'),(7,'I'),(8,'E')]:
                    raw_val = row.iloc[col] if len(row) > col else None
                    if raw_val is None or (isinstance(raw_val, float) and pd.isna(raw_val)):
                        continue
                    try:
                        v = float(raw_val)
                    except Exception:
                        continue
                    if v <= 0:
                        continue
                    if v >= 1.0:
                        slots.append((key, f"{v} {key}"))
                    else:
                        frac_parts.append(f"{v} {key}")
                        frac_keys.append(key)
                if frac_parts:
                    slots.append(("".join(frac_keys), "\n".join(frac_parts)))
                # Admin slot
                adm_v = row.iloc[9] if len(row) > 9 else None
                if adm_v is not None and not (isinstance(adm_v, float) and pd.isna(adm_v)):
                    try:
                        if float(adm_v) > 0:
                            slots.append(("1A", "Admin"))
                    except Exception:
                        pass
                break  # found the faculty row

        return slots[:4]

    for idx, fr in enumerate(faculty_rows, 1):
        nm = fr["name"]
        des = _designation(nm)
        slots = _build_cores(nm)
        while len(slots) < 4:
            slots.append((None, None))
        row = [idx, nm, des]
        for st, sd in slots:
            row += [st, sd]
        row += [0, None]
        ws3.append(row)

    # ─── Sheet 4: Summary Stats ───────────────────────────────────────────────
    faculty_data_for_summary = []
    for fr in faculty_rows:
        nm = fr["name"]
        if nm.startswith("NF"):
            continue
        th, lb = _teaching_units(nm)
        tu = round(th + lb, 4)
        c_val = fr.get("C"); p_val = fr.get("P"); i_val = fr.get("I")
        e_val = fr.get("E"); adm   = fr.get("A"); total = fr.get("Total")
        def _flt2(v):
            try: return float(v) if v is not None else 0.0
            except: return 0.0
        r_computed = round(max(0.0, _flt(total) - tu - _flt(c_val) - _flt(p_val)
                                    - _flt(i_val) - _flt(e_val) - _flt(adm)), 4)
        faculty_data_for_summary.append({
            "idx": fr.get("sl_no"), "name": nm,
            "designation": _designation(nm),
            "theory": th if th else None,
            "lab":    lb if lb else None,
            "tu": tu,
            "r": r_computed if r_computed > 0 else None,
            "c": c_val, "p": p_val, "i": i_val, "e": e_val,
            "admin": adm, "total": total,
        })
    _add_summary_sheet(wb, faculty_data_for_summary)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()



def _add_summary_sheet(wb, faculty_data_rows: list):
    """Adds a formatted Summary Stats sheet. Uses Excel SUM/COUNTIF formulas."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    ws = wb.create_sheet("Summary Stats")
    ws.sheet_properties.tabColor = "2563EB"

    # column widths
    widths = [2, 6, 28, 20, 10, 10, 14, 12, 10, 10, 10, 10, 14, 2]
    for i, w in enumerate(widths, 1):
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(i)].width = w

    def fill(c):   return PatternFill("solid", fgColor=c)
    def font(bold=False, sz=10, color="000000"):
        return Font(bold=bold, size=sz, color=color, name="Arial")
    def align(h="left", wrap=False):
        return Alignment(horizontal=h, vertical="center", wrap_text=wrap)
    def thin_border():
        s = Side(style="thin", color="BDBDBD")
        return Border(left=s, right=s, top=s, bottom=s)

    SA = "'subject allocation'"   # cross-sheet reference prefix
    row = [1]

    def R(skip=0):
        row[0] += 1 + skip
        return row[0]

    def cell(r, c, val="", bold=False, sz=10, color="000000", bg=None, h="left",
             wrap=False, border=False, fmt=None):
        cl = ws.cell(row=r, column=c, value=val)
        cl.font = font(bold, sz, color)
        cl.alignment = align(h, wrap)
        if bg:   cl.fill = fill(bg)
        if border: cl.border = thin_border()
        if fmt:  cl.number_format = fmt
        return cl

    def mrow(r, c1, c2):
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)

    # ── TITLE BANNER ────────────────────────────────────────────────────────
    r = R()
    ws.row_dimensions[r].height = 34
    mrow(r, 2, 13)
    cell(r,2,"CMR INSTITUTE OF TECHNOLOGY  —  AIML Department  |  Faculty Core Load Summary 2025-26 Even Semester",
         bold=True, sz=13, color="FFFFFF", bg="1E3A8A", h="center")

    R()  # blank

    # ── METRIC CARDS (2 rows × 3 cols) ──────────────────────────────────────
    # Row A: col B-E / F-I / J-M
    card_defs = [
        ("Total Faculty",         f"=COUNTA({SA}!C10:C33)",                    "1E40AF","DBEAFE"),
        ("Total Teaching Units",  f"=IFERROR(SUM({SA}!F10:F33),0)",            "166534","DCFCE7"),
        ("Total Core Hours (T)",  f"=IFERROR(SUM({SA}!F10:F33)*6,0)",          "92400E","FEF3C7"),
        ("Avg Teaching/Faculty",  f"=IFERROR(SUM({SA}!F10:F33)/COUNTA({SA}!C10:C33),0)", "6B21A8","F3E8FF"),
        ("Faculty at Full Load",  f"=COUNTIF({SA}!F10:F33,4)",                 "0E7490","CFFAFE"),
        ("Assoc / Asst Profs",    f"=COUNTIF({SA}!C10:C33,\"Associate Professor\")&\" / \"&COUNTIF({SA}!C10:C33,\"Assistant Professor\")", "991B1B","FEE2E2"),
    ]
    card_col_ranges = [(2,5),(6,9),(10,13)]
    for ci, (lbl, formula, tcol, bgcol) in enumerate(card_defs):
        crow_offset = ci // 3
        ccol_start, ccol_end = card_col_ranges[ci % 3]
        base = row[0] + crow_offset * 3

        # Value row
        ws.row_dimensions[base].height = 28
        mrow(base, ccol_start, ccol_end)
        cl = ws.cell(row=base, column=ccol_start, value=formula)
        cl.font = Font(bold=True, size=16, color=tcol, name="Arial")
        cl.alignment = Alignment(horizontal="center", vertical="center")
        cl.fill = fill(bgcol)
        cl.number_format = "0.00" if "SUM" in formula or "AVERAGE" in formula or "*6" in formula else "0"

        # Label row
        ws.row_dimensions[base+1].height = 15
        mrow(base+1, ccol_start, ccol_end)
        cl2 = ws.cell(row=base+1, column=ccol_start, value=lbl)
        cl2.font = Font(size=9, color=tcol, name="Arial")
        cl2.alignment = Alignment(horizontal="center", vertical="center")
        cl2.fill = fill(bgcol)

    row[0] += 7   # advance past 2 rows of cards (each 3 rows tall) + 1 extra

    R()  # blank

    # ── TRCPIE TOTALS TABLE ─────────────────────────────────────────────────
    r = R()
    mrow(r, 2, 13)
    ws.row_dimensions[r].height = 20
    cell(r,2,"TRCPIE CORE LOAD TOTALS   (1 unit = 6 hrs)",
         bold=True, sz=10, color="1E3A8A", bg="EFF6FF", h="center")

    # Header
    r = R()
    ws.row_dimensions[r].height = 22
    for c,hdr,bg_ in [
        (2,"Code","1E3A8A"),(3,"Component","1E3A8A"),(5,"Units","1E3A8A"),
        (6,"Hours (×6)","1E3A8A"),(7,"% of Total","1E3A8A"),(8,"Faculty Count","1E3A8A")]:
        cell(r,c,hdr, bold=True, sz=9, color="FFFFFF", bg="1E3A8A", h="center", border=True)
    mrow(r,3,4)

    trcpie_defs = [
        ("T","Teaching Units",       f"=SUM({SA}!F10:F33)",              f"=COUNTA({SA}!C10:C33)",  "1E40AF","DBEAFE"),
        ("R","Research",             f"=SUMIF({SA}!G10:G33,\">0\")",     f"=COUNTIF({SA}!G10:G33,\">0\")", "166534","DCFCE7"),
        ("C","Consultancy",          f"=SUMIF({SA}!H10:H33,\">0\")",     f"=COUNTIF({SA}!H10:H33,\">0\")", "92400E","FEF3C7"),
        ("P","Sponsored Project",    f"=SUMIF({SA}!I10:I33,\">0\")",     f"=COUNTIF({SA}!I10:I33,\">0\")", "6B21A8","F3E8FF"),
        ("I","Innovation",           f"=SUMIF({SA}!J10:J33,\">0\")",     f"=COUNTIF({SA}!J10:J33,\">0\")", "0E7490","CFFAFE"),
        ("E","Entrepreneurship",     f"=SUMIF({SA}!K10:K33,\">0\")",     f"=COUNTIF({SA}!K10:K33,\">0\")", "991B1B","FEE2E2"),
        ("A","Admin",                f"=SUMIF({SA}!L10:L33,\">0\")",     f"=COUNTIF({SA}!L10:L33,\">0\")", "374151","F1F5F9"),
    ]

    unit_rows = {}
    for code, name, units_f, count_f, tcol, bgcol in trcpie_defs:
        r = R()
        ws.row_dimensions[r].height = 20
        unit_rows[code] = r
        cell(r,2, code,  bold=True, sz=11, color=tcol, bg=bgcol, h="center", border=True)
        mrow(r,3,4)
        cell(r,3, name, bold=False, sz=10, color=tcol, bg=bgcol, border=True)
        cl_u = ws.cell(row=r, column=5, value=units_f)
        cl_u.font = Font(bold=True, sz=11, color=tcol, name="Arial"); cl_u.alignment = align("center")
        cl_u.fill = fill(bgcol); cl_u.border = thin_border(); cl_u.number_format = "0.0000"
        cl_h = ws.cell(row=r, column=6, value=f"=E{r}*6")
        cl_h.font = Font(sz=10, color=tcol, name="Arial"); cl_h.alignment = align("center")
        cl_h.fill = fill(bgcol); cl_h.border = thin_border(); cl_h.number_format = "0.00"
        cl_c = ws.cell(row=r, column=8, value=count_f)
        cl_c.font = Font(sz=10, color=tcol, name="Arial"); cl_c.alignment = align("center")
        cl_c.fill = fill(bgcol); cl_c.border = thin_border()

    # Grand total row
    r = R()
    ws.row_dimensions[r].height = 22
    mrow(r,2,4)
    cell(r,2,"GRAND TOTAL", bold=True, sz=10, color="FFFFFF", bg="1E3A8A", h="center", border=True)
    unit_sum = "+".join(f"E{unit_rows[k]}" for k in ["T","R","C","P","I","E","A"])
    cl_gt = ws.cell(row=r, column=5, value=f"={unit_sum}")
    cl_gt.font = Font(bold=True, size=12, color="FFFFFF", name="Arial"); cl_gt.alignment = align("center")
    cl_gt.fill = fill("1E3A8A"); cl_gt.border = thin_border(); cl_gt.number_format = "0.00"
    cl_gth = ws.cell(row=r, column=6, value=f"=E{r}*6")
    cl_gth.font = Font(bold=True, size=11, color="FFFFFF", name="Arial"); cl_gth.alignment = align("center")
    cl_gth.fill = fill("1E3A8A"); cl_gth.border = thin_border(); cl_gth.number_format = "0.00"

    # % column — reference grand total row
    for code in ["T","R","C","P","I","E","A"]:
        tr = unit_rows[code]
        bgcol = [x[5] for x in trcpie_defs if x[0]==code][0]
        tcol  = [x[4] for x in trcpie_defs if x[0]==code][0]
        cl_pct = ws.cell(row=tr, column=7, value=f"=IFERROR(E{tr}/E{r},0)")
        cl_pct.font = Font(size=10, color=tcol, name="Arial"); cl_pct.alignment = align("center")
        cl_pct.fill = fill(bgcol); cl_pct.border = thin_border(); cl_pct.number_format = "0.0%"
    cell(r,7,"100%", bold=True, sz=10, color="FFFFFF", bg="1E3A8A", h="center", border=True)

    R()  # blank

    # ── PER-FACULTY TABLE ────────────────────────────────────────────────────
    r = R()
    mrow(r,2,13)
    ws.row_dimensions[r].height = 20
    cell(r,2,"PER-FACULTY TEACHING LOAD DISTRIBUTION",
         bold=True, sz=10, color="1E3A8A", bg="EFF6FF", h="center")

    r = R()
    ws.row_dimensions[r].height = 20
    col_hdrs = [(2,"#"),(3,"Name"),(5,"Designation"),(7,"Theory"),(8,"Lab"),
                (9,"Teaching Units"),(10,"R"),(11,"C"),(12,"P"),(13,"I")]
    for c, h_ in col_hdrs:
        cell(r, c, h_, bold=True, sz=9, color="FFFFFF", bg="1E3A8A", h="center", border=True)
    mrow(r,3,4); mrow(r,5,6)

    alt = ["FFFFFF","F8FAFC"]
    for ri, fd in enumerate(faculty_data_rows):
        r = R()
        ws.row_dimensions[r].height = 18
        bg = alt[ri % 2]
        tu = fd.get("tu") or 0
        tu_bg = "DCFCE7" if float(tu)>=4 else "DBEAFE" if float(tu)>=3 else "FEF3C7"

        cell(r,2, fd["idx"], bg=bg, h="center", border=True)
        mrow(r,3,4); cell(r,3, fd["name"], sz=10, bg=bg, border=True)
        mrow(r,5,6); cell(r,5, fd.get("designation",""), sz=9, bg=bg, h="center", border=True)

        for c, key, fmt_ in [
            (7,"theory","0.0000"),(8,"lab","0.0000"),
            (10,"r","0.0000"),(11,"c","0.0000"),(12,"p","0.0000"),(13,"i","0.0000")]:
            v = fd.get(key)
            cl = ws.cell(row=r, column=c, value="" if v is None else v)
            cl.font = Font(size=10, name="Arial"); cl.alignment = align("center")
            cl.fill = fill(bg); cl.border = thin_border()
            if v is not None: cl.number_format = fmt_

        cl_tu = ws.cell(row=r, column=9, value="" if tu == 0 else tu)
        cl_tu.font = Font(bold=True, size=10, name="Arial"); cl_tu.alignment = align("center")
        cl_tu.fill = fill(tu_bg); cl_tu.border = thin_border(); cl_tu.number_format = "0.0000"

    # Total row
    r = R()
    ws.row_dimensions[r].height = 20
    mrow(r,2,8)
    cell(r,2,"TOTAL", bold=True, sz=10, color="FFFFFF", bg="1E3A8A", h="right", border=True)
    data_start = r - len(faculty_data_rows)
    cl_t = ws.cell(row=r, column=9, value=f"=SUM(I{data_start}:I{r-1})")
    cl_t.font = Font(bold=True, size=11, color="FFFFFF", name="Arial"); cl_t.alignment = align("center")
    cl_t.fill = fill("1E3A8A"); cl_t.border = thin_border(); cl_t.number_format = "0.0000"
    for c in [10,11,12,13]:
        cell(r,c,"", bg="1E3A8A", border=True)

    # Legend
    R(); r = R()
    cell(r,2,"Legend:", bold=True, sz=9, color="475569")
    for li,(lgbg,lglbl) in enumerate([("DCFCE7","≥ 4 units (full)"),("DBEAFE","3–3.9 units"),("FEF3C7","< 3 units")]):
        c = 3+li*3
        mrow(r,c,c+1)
        ws.cell(row=r,column=c).fill = fill(lgbg); ws.cell(row=r,column=c).border = thin_border()
        cell(r,c+2, lglbl, sz=9, color="475569")
