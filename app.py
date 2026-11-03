"""
Attendance Register Filler  (Streamlit)

Takes:
  1) A "required" file that has, per student, the number of classes attended.
  2) An attendance register with one column per class date (P / A cells).

For every student matched by enrollment number, the app marks exactly
(total classes - attended) randomly chosen dates as Absent (highlighted red)
and the rest as Present, then updates the total columns.

Run:
    pip install streamlit openpyxl pandas
    streamlit run app.py
"""

import io
import random
import re

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string

st.set_page_config(page_title="Attendance Register Filler", page_icon="📋", layout="wide")


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def norm(v):
    """Normalise an enrollment value for matching."""
    if v is None:
        return ""
    return re.sub(r"\s+", "", str(v)).upper()


def load_wb(file_bytes):
    return openpyxl.load_workbook(io.BytesIO(file_bytes))


def sheet_preview(ws, max_rows=8, max_cols=30):
    rows = []
    for r in ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_rows),
                          max_col=min(ws.max_column, max_cols), values_only=True):
        rows.append(["" if v is None else str(v) for v in r])
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df.columns = [get_column_letter(i + 1) for i in range(df.shape[1])]
    df.index = range(1, len(df) + 1)
    return df


def find_header_col(ws, header_row, keyword):
    """Find first column whose header contains keyword (case-insensitive)."""
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None and keyword.lower() in str(v).lower():
            return c
    return None


def guess_register_layout(ws, header_row):
    """Guess enrollment column, date columns and total columns from headers."""
    enroll = (find_header_col(ws, header_row, "regno")
              or find_header_col(ws, header_row, "enroll")
              or find_header_col(ws, header_row, "reg"))
    c_total = find_header_col(ws, header_row, "total_lectures") or find_header_col(ws, header_row, "total")
    c_att = find_header_col(ws, header_row, "attended")
    c_pct = find_header_col(ws, header_row, "percent")

    date_re = re.compile(r"\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}")
    date_cols = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None and date_re.search(str(v)):
            date_cols.append(c)
    first = min(date_cols) if date_cols else None
    last = max(date_cols) if date_cols else None
    return dict(enroll=enroll, first=first, last=last,
                total=c_total, attended=c_att, percent=c_pct)


def guess_required_cols(ws, header_row, register_ids):
    """Enrollment column = column with the most overlap with register ids.
    Attended column = header containing 'attended', else first numeric column after enrollment."""
    best_col, best_hits = None, 0
    for c in range(1, ws.max_column + 1):
        hits = 0
        for r in range(header_row + 1, ws.max_row + 1):
            if norm(ws.cell(r, c).value) in register_ids:
                hits += 1
        if hits > best_hits:
            best_col, best_hits = c, hits

    att_col = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None and "attended" in str(v).lower():
            att_col = c
            break
    if att_col is None and best_col is not None:
        for c in range(best_col + 1, ws.max_column + 1):
            vals = [ws.cell(r, c).value for r in range(header_row + 1, min(ws.max_row, header_row + 15) + 1)]
            vals = [v for v in vals if v is not None]
            if vals and all(isinstance(v, (int, float)) for v in vals):
                att_col = c
                break
    return best_col, att_col


def process(req_bytes, reg_bytes, cfg):
    """Core logic. Returns (output_bytes, report_df, summary_dict)."""
    req_ws = load_wb(req_bytes)[cfg["req_sheet"]]
    reg_wb = load_wb(reg_bytes)
    reg_ws = reg_wb[cfg["reg_sheet"]]

    # --- read required attended counts ---
    req = {}
    req_names = {}
    for r in range(cfg["req_header_row"] + 1, req_ws.max_row + 1):
        key = norm(req_ws.cell(r, cfg["req_enroll_col"]).value)
        val = req_ws.cell(r, cfg["req_att_col"]).value
        if not key:
            continue
        try:
            val = int(round(float(val)))
        except (TypeError, ValueError):
            continue
        req[key] = val
        if cfg.get("req_name_col"):
            req_names[key] = req_ws.cell(r, cfg["req_name_col"]).value

    rng = random.Random(cfg["seed"]) if cfg["seed"] is not None else random.Random()
    red = PatternFill("solid", start_color="FFFF0000", end_color="FFFF0000")
    no_fill = PatternFill(fill_type=None)

    first, last = cfg["first_date_col"], cfg["last_date_col"]
    date_cols = list(range(first, last + 1))
    total_classes = cfg["total_classes"]
    if total_classes > len(date_cols):
        raise ValueError(
            f"Total classes ({total_classes}) is more than the number of date columns ({len(date_cols)})."
        )

    report = []
    matched_keys = set()
    for r in range(cfg["reg_header_row"] + 1, reg_ws.max_row + 1):
        raw = reg_ws.cell(r, cfg["reg_enroll_col"]).value
        key = norm(raw)
        if not key:
            continue
        if key not in req:
            report.append(dict(Row=r, Enrollment=raw, Attended="", Absent="", Status="Not in required file - left unchanged"))
            continue

        matched_keys.add(key)
        attended = req[key]
        if attended < 0 or attended > total_classes:
            report.append(dict(Row=r, Enrollment=raw, Attended=attended, Absent="",
                               Status=f"Skipped - attended count invalid (must be 0..{total_classes})"))
            continue

        absent_n = total_classes - attended
        # classes counted = first `total_classes` date columns; extra columns (if any) untouched
        counted_cols = date_cols[:total_classes]
        absent_cols = set(rng.sample(counted_cols, absent_n))
        for c in counted_cols:
            cell = reg_ws.cell(r, c)
            if c in absent_cols:
                cell.value = cfg["absent_label"]
                if cfg["highlight"]:
                    cell.fill = red
            else:
                cell.value = cfg["present_label"]
                cell.fill = no_fill

        if cfg["update_totals"]:
            if cfg["col_total"]:
                reg_ws.cell(r, cfg["col_total"]).value = total_classes
            if cfg["col_attended"]:
                reg_ws.cell(r, cfg["col_attended"]).value = attended
            if cfg["col_percent"]:
                pct = attended / total_classes * 100 if total_classes else 0
                reg_ws.cell(r, cfg["col_percent"]).value = (
                    int(pct) if cfg["pct_mode"] == "Whole number (rounded down)" else round(pct, 2)
                )

        report.append(dict(Row=r, Enrollment=raw, Attended=attended, Absent=absent_n, Status="Updated"))

    for key, val in req.items():
        if key not in matched_keys:
            report.append(dict(Row="-", Enrollment=key, Attended=val, Absent="",
                               Status="In required file but not in register"))

    buf = io.BytesIO()
    reg_wb.save(buf)
    df = pd.DataFrame(report)
    if not df.empty:
        df["Row"] = df["Row"].map(str)
        df["Enrollment"] = df["Enrollment"].map(lambda v: "" if v is None else str(v))
        df["Attended"] = pd.to_numeric(df["Attended"], errors="coerce").astype("Int64")
        df["Absent"] = pd.to_numeric(df["Absent"], errors="coerce").astype("Int64")
    summary = dict(
        updated=int((df["Status"] == "Updated").sum()) if not df.empty else 0,
        not_in_required=int(df["Status"].str.startswith("Not in required").sum()) if not df.empty else 0,
        not_in_register=int(df["Status"].str.startswith("In required").sum()) if not df.empty else 0,
        skipped=int(df["Status"].str.startswith("Skipped").sum()) if not df.empty else 0,
    )
    return buf.getvalue(), df, summary


def letter_options(n):
    return [get_column_letter(i) for i in range(1, n + 1)]


def idx_of(letter_list, col_idx, default=0):
    if not letter_list:
        return 0
    default = max(0, min(default, len(letter_list) - 1))
    if col_idx is None:
        return default
    i = col_idx - 1
    return i if 0 <= i < len(letter_list) else default


# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------
st.title("📋 Attendance Register Filler")
st.write(
    "Upload the file that holds **classes attended** per student and the **attendance register** "
    "(one column per class date). The app marks the right number of random dates as **Absent** "
    "(red) so each student's total matches."
)

c1, c2 = st.columns(2)
with c1:
    req_file = st.file_uploader("1) Attendance Required file (has classes attended)", type=["xlsx", "xlsm"], key="req")
with c2:
    reg_file = st.file_uploader("2) Attendance Register file (to be updated)", type=["xlsx", "xlsm"], key="reg")

if not (req_file and reg_file):
    st.info("Upload both files to continue.")
    st.stop()

req_bytes, reg_bytes = req_file.getvalue(), reg_file.getvalue()
req_wb, reg_wb_probe = load_wb(req_bytes), load_wb(reg_bytes)

st.divider()
st.subheader("Step 1 - Register layout")
rc1, rc2 = st.columns([1, 3])
with rc1:
    reg_sheet = st.selectbox("Register sheet", reg_wb_probe.sheetnames)
    reg_ws = reg_wb_probe[reg_sheet]
    reg_header_row = st.number_input("Register header row", min_value=1, value=1, step=1)
with rc2:
    st.caption("Register preview")
    st.dataframe(sheet_preview(reg_ws), width="stretch")

g = guess_register_layout(reg_ws, reg_header_row)
letters_reg = letter_options(reg_ws.max_column)
if g["first"] is None:
    st.warning("Could not auto-detect date columns from the header row. Pick them manually below.")

a, b, c = st.columns(3)
with a:
    reg_enroll_col = column_index_from_string(
        st.selectbox("Enrollment column (register)", letters_reg, index=idx_of(letters_reg, g["enroll"]))
    )
with b:
    first_date_col = column_index_from_string(
        st.selectbox("First date column", letters_reg, index=idx_of(letters_reg, g["first"], 3))
    )
with c:
    last_date_col = column_index_from_string(
        st.selectbox("Last date column", letters_reg, index=idx_of(letters_reg, g["last"], 3))
    )
if last_date_col < first_date_col:
    st.error("Last date column must be at or after the first date column.")
    st.stop()
n_date_cols = last_date_col - first_date_col + 1

st.divider()
st.subheader("Step 2 - Required (attended) file")
qc1, qc2 = st.columns([1, 3])
with qc1:
    req_sheet = st.selectbox("Required sheet", req_wb.sheetnames)
    req_ws = req_wb[req_sheet]
    req_header_row = st.number_input("Required header row", min_value=0, value=1, step=1,
                                     help="Rows after this are treated as data. Use 0 if there is no header.")
with qc2:
    st.caption("Required file preview")
    st.dataframe(sheet_preview(req_ws, max_rows=10), width="stretch")

register_ids = {
    norm(reg_ws.cell(r, reg_enroll_col).value)
    for r in range(reg_header_row + 1, reg_ws.max_row + 1)
    if reg_ws.cell(r, reg_enroll_col).value
}
g_enroll, g_att = guess_required_cols(req_ws, req_header_row, register_ids)
letters_req = letter_options(req_ws.max_column)

d, e, f = st.columns(3)
with d:
    req_enroll_col = column_index_from_string(
        st.selectbox("Enrollment column (required)", letters_req, index=idx_of(letters_req, g_enroll))
    )
with e:
    req_att_col = column_index_from_string(
        st.selectbox("Classes attended column", letters_req, index=idx_of(letters_req, g_att))
    )
with f:
    name_opts = ["(none)"] + letters_req
    name_sel = st.selectbox("Name column (optional)", name_opts, index=0)
    req_name_col = None if name_sel == "(none)" else column_index_from_string(name_sel)

st.divider()
st.subheader("Step 3 - Options")
o1, o2, o3 = st.columns(3)
with o1:
    total_classes = st.number_input("Total classes", min_value=1, max_value=n_date_cols, value=n_date_cols, step=1,
                                    help=f"Detected {n_date_cols} date columns in the selected range.")
    present_label = st.text_input("Present label", "P")
    absent_label = st.text_input("Absent label", "A")
with o2:
    highlight = st.checkbox("Highlight absent cells in red", value=True)
    update_totals = st.checkbox("Update total columns", value=True)
    use_seed = st.checkbox("Use fixed random seed (repeatable result)", value=False)
    seed = st.number_input("Seed", value=42, step=1, disabled=not use_seed)
with o3:
    letters_tot = ["(none)"] + letters_reg
    def tot_idx(col):
        return col if col else 0  # +1 for '(none)' offset, -1 for 1-based => col
    col_total_sel = st.selectbox("Total classes column", letters_tot, index=tot_idx(g["total"]), disabled=not update_totals)
    col_att_sel = st.selectbox("Total attended column", letters_tot, index=tot_idx(g["attended"]), disabled=not update_totals)
    col_pct_sel = st.selectbox("Percentage column", letters_tot, index=tot_idx(g["percent"]), disabled=not update_totals)
    pct_mode = st.radio("Percentage format", ["Whole number (rounded down)", "2 decimals"], horizontal=True,
                        disabled=not update_totals)


def to_col(sel):
    return None if sel == "(none)" else column_index_from_string(sel)


cfg = dict(
    req_sheet=req_sheet, reg_sheet=reg_sheet,
    req_header_row=int(req_header_row), reg_header_row=int(reg_header_row),
    req_enroll_col=req_enroll_col, req_att_col=req_att_col, req_name_col=req_name_col,
    reg_enroll_col=reg_enroll_col, first_date_col=first_date_col, last_date_col=last_date_col,
    total_classes=int(total_classes), present_label=present_label, absent_label=absent_label,
    highlight=highlight, update_totals=update_totals,
    col_total=to_col(col_total_sel), col_attended=to_col(col_att_sel), col_percent=to_col(col_pct_sel),
    pct_mode=pct_mode, seed=int(seed) if use_seed else None,
)

st.divider()
if st.button("▶ Process", type="primary"):
    try:
        out_bytes, report_df, summary = process(req_bytes, reg_bytes, cfg)
        st.session_state["result"] = (out_bytes, report_df, summary)
    except Exception as ex:  # noqa: BLE001
        st.session_state.pop("result", None)
        st.error(f"Could not process: {ex}")

if "result" in st.session_state:
    out_bytes, report_df, summary = st.session_state["result"]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Updated", summary["updated"])
    m2.metric("In register, not in required", summary["not_in_required"])
    m3.metric("In required, not in register", summary["not_in_register"])
    m4.metric("Skipped (invalid count)", summary["skipped"])

    out_name = re.sub(r"\.xlsm?$", "", reg_file.name, flags=re.I) + "_updated.xlsx"
    st.download_button(
        "⬇ Download updated register",
        data=out_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.subheader("Report")
    only_issues = st.checkbox("Show only rows that were not updated", value=False)
    show = report_df[report_df["Status"] != "Updated"] if only_issues else report_df
    st.dataframe(show, width="stretch", hide_index=True)
    st.caption("Rows not matched by enrollment number are left exactly as they were in the register.")
