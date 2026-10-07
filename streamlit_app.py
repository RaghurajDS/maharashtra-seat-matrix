"""
Streamlit app: upload the Maharashtra CET seat-matrix PDF -> download Excel.

Run:
    pip install streamlit pdfplumber openpyxl pandas
    streamlit run streamlit_app.py

Keep this file in the SAME folder as seat_matrix_pdf_to_excel.py
"""
import io
import contextlib

import pandas as pd
import streamlit as st

import seat_matrix_pdf_to_excel as conv

st.set_page_config(page_title="Maharashtra Seat Matrix → Excel", page_icon="📄", layout="wide")
st.title("📄 Maharashtra Seat Matrix → Excel")
st.caption("Upload the State CET Cell 'Provisional Seat Matrix for MBBS/BDS' PDF and download a checked Excel file.")

# ---------------- sidebar options ----------------
with st.sidebar:
    st.header("Output")
    fmt = st.radio(
        "Excel format",
        ["One row per seat entry (long)", "Same layout as the PDF (+ Checks sheet)"],
        help="Long format: Sr, Code, College, Aided, Min, GOI, row type, Category, Seats.",
    )
    long_fmt = fmt.startswith("One row")
    keep_zeros = section_col = False
    if long_fmt:
        keep_zeros = st.checkbox("Keep seats printed as 0", value=False)
        section_col = st.checkbox("Add 'Section' column (Govt / Pvt / Pvt Univ)", value=False)

pdf_file = st.file_uploader("Seat matrix PDF", type=["pdf"])

if pdf_file is None:
    st.info("Choose a PDF to begin.")
    st.stop()

# ---------------- convert ----------------
with st.spinner("Reading PDF…"):
    try:
        sections = conv.parse_pdf(io.BytesIO(pdf_file.getvalue()))
    except Exception as exc:                                    # unreadable / not a PDF
        st.error(f"Could not read this PDF: {exc}")
        st.stop()

sections = [s for s in sections if s["colleges"] and s["total"]]
if not sections:
    st.error(
        "No Maharashtra CET seat-matrix tables (GOVT./PVT. COLLEGES: MBBS/BDS) were found in this PDF. "
        "This app only reads that specific report."
    )
    st.stop()

# validation (the script prints its report; capture it)
log = io.StringIO()
with contextlib.redirect_stdout(log):
    problems = conv.validate(sections)

# ---------------- results ----------------
c1, c2, c3 = st.columns(3)
c1.metric("Sections found", len(sections))
c2.metric("Colleges", sum(len(s["colleges"]) for s in sections))
c3.metric("Check problems", problems)

if problems == 0:
    st.success("All checks passed: row totals, college totals, printed Total rows and Req. Of all reconcile.")
else:
    st.warning("Some checks failed – see the report below before using the file.")
with st.expander("Validation report", expanded=problems != 0):
    st.code(log.getvalue() or "No output", language="text")

# build the workbook in memory
buf = io.BytesIO()
if long_fmt:
    with contextlib.redirect_stdout(io.StringIO()):
        conv.build_long_workbook(sections, buf, keep_zeros, section_col)
else:
    conv.build_workbook(sections, buf)
buf.seek(0)

base = pdf_file.name.rsplit(".", 1)[0]
st.download_button(
    "⬇️ Download Excel",
    data=buf.getvalue(),
    file_name=f"{base}_{'long' if long_fmt else 'layout'}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

# ---------------- preview ----------------
st.subheader("Preview")
if long_fmt:
    cols = ["Sr", "Code", "College", "Aided", "Min", "GOI", "Type", "Category", "Seats", "Section"]
    courses = {}
    for s in sections:
        courses.setdefault(s["name"].split(":")[1].strip(), []).extend(conv.long_rows(s, keep_zeros))
    tabs = st.tabs([f"{c} ({len(r)} rows)" for c, r in courses.items()])
    for tab, (course, rows) in zip(tabs, courses.items()):
        with tab:
            df = pd.DataFrame(rows, columns=cols)
            if not section_col:
                df = df.drop(columns="Section")
            st.dataframe(df, hide_index=True)
else:
    rows = [(s["name"], len(s["colleges"]), s["req_text"]) for s in sections]
    st.dataframe(pd.DataFrame(rows, columns=["Section", "Colleges", "Req. Of"]), hide_index=True)
    st.caption("The downloaded workbook has one sheet per section plus a 'Checks' sheet of live formulas.")
