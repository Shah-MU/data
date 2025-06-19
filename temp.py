"""
tables.py ───────────────────────────────────────────────────────────────
Reusable subtotal-table builder for Streamlit + st-aggrid, with Excel
download that preserves styling and filtering.

Run:
    pip install streamlit streamlit-aggrid xlsxwriter
    streamlit run tables.py
"""


import io, json, streamlit as st, pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, GridUpdateMode
from xlsxwriter.utility import xl_col_to_name

st.set_page_config(layout="wide")

ROW_ORDER = [
    "Term & Prime Loans to  Financials",
    "Term & Prime Loans to Non Fiancials",
    "iVision Loans",
    "Non Performing loans-net of provision",
    "Total Lending book",
    "HQLA Level1-CB Eligibile - CCP Eligible - Unencumbered",
    "HQLA Level1-CB Eligibile - Non CCP Eligible - Unencumbered",
    "HQLA Level2A/15 - Unencumbered",
    "HQLA Level2B/25 - Unencumbered",
    "HQLA Level2B/50 - Unencumbered",
    "Non HQLA Securities",
    "HQLA Encumbered LVTS",
    "HQLA Encumbered others",
    "Total MTM value",
    "Rev Repos with HQLA Level1- Intragroup",
    "Rev Repos with HQLA Level1- Non Intragroup",
    "Rev Repos with non-HQLA",
    "Others",
    "Secured Lending",
    "Central Bank Term Loan",
    "Interbank Term Loan",
    "Interbank Lending",
    "Intra-Treasuries (Paris)",
    "Intra-Treasuries (arbitage)",
    "Other Affliates (inc Leasing)",
    "Intragroup Lending",
    "Intragroup Lending - Nostros Balances",
    "Intragroup Lending - Overdraft Loro's/DDA's",
    "Intergroup Lending",
    "FBBD",
    "Other assets from Finance B/S",
    "Total other assets",
    "Total Assets",
]

TENOR_COLS = ["O/N", "1D-1W", "1-2W", "2W-1M", "1-3M", "3-12M", ">1 Y"]
LCR_COLS   = ["LCR CF 30D", "W.LCR CF"]
DF_COLS    = ["Total", *TENOR_COLS, *LCR_COLS, "CF"]


# ────────────────────────────── editable-JS helper ─────────────────────
def _editable_js(locked_cols: list[str], editable_row: str) -> JsCode:
    locked_json   = json.dumps(locked_cols)
    row_condition = (
        f"p.data.Metric === {json.dumps(editable_row)}" if editable_row else "false"
    )
    return JsCode(f"""
function(p){{
  const f = p.colDef.field;
  if (f === 'CF')                     return true;
  if ({locked_json}.includes(f))      return false;
  return {row_condition};
}}""")


# ────────────────────────────── Excel export helper ────────────────────
def _write_excel(
    df: pd.DataFrame,
    filename: str,
    *,
    subtotal_rows: list[str],
    grand_total_row: str,
) -> io.BytesIO:
    """
    Create an Excel file in-memory with colouring, filters and frozen header.
    Values of ``pd.NA`` are converted to ``None`` to avoid xlsxwriter TypeError.
    """
    # Prepare a DataFrame fit for Excel (no index, pd.NA → None)
    df_reset = df.reset_index(names="Metric")
    df_excel = (df_reset.astype(object)  # keep numbers but allow None
                         .where(pd.notna(df_reset), None))

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df_excel.to_excel(writer, sheet_name="Data", index=False)
        wb  = writer.book
        ws  = writer.sheets["Data"]

        # -------- column formats --------
        fmt_sub   = wb.add_format({"bg_color": "#d4f7d4", "bold": True})
        fmt_total = wb.add_format({"bg_color": "#cce5ff", "bold": True})
        fmt_lcr   = wb.add_format({"bg_color": "#ffe6f2"})
        fmt_cf    = wb.add_format({"bg_color": "#f5f5f5"})
        fmt_red   = wb.add_format({"bg_color": "#ffcccc"})
        fmt_green = wb.add_format({"bg_color": "#d4f7d4"})

        # Freeze header and enable auto-filter
        ws.freeze_panes(1, 1)
        ws.autofilter(0, 0, df_excel.shape[0], df_excel.shape[1] - 1)

        # Row-wise background (subtotals & grand total)
        for r, metric in enumerate(df_excel["Metric"], start=1):  # +1 for header
            if metric in subtotal_rows:
                ws.set_row(r, None, fmt_sub)
            elif metric == grand_total_row:
                ws.set_row(r, None, fmt_total)

        # Column-wise background (LCR & CF)
        for col in LCR_COLS:
            idx = df_excel.columns.get_loc(col)
            ws.set_column(idx, idx, None, fmt_lcr)
        cf_idx = df_excel.columns.get_loc("CF")
        ws.set_column(cf_idx, cf_idx, None, fmt_cf)

        # Conditional formatting on the "Total" column
        total_idx  = df_excel.columns.get_loc("Total")
        ws.conditional_format(
            1, total_idx, df_excel.shape[0], total_idx,
            {"type": "cell", "criteria": "<", "value": 0, "format": fmt_red},
        )
        ws.conditional_format(
            1, total_idx, df_excel.shape[0], total_idx,
            {"type": "cell", "criteria": ">=", "value": 0, "format": fmt_green},
        )

    buf.seek(0)
    return buf


# ───────────────────────────── subtotal-table builder ──────────────────
def build_subtotal_table(
    *,
    title: str,
    row_order: list[str],
    subtotal_rows: list[str],
    grand_total_row: str,
    preload_data: dict[str, list | dict] | None = None,
    editable_row: str = "",
) -> None:
    """Render an AG-Grid + matching Excel download button."""
    df = pd.DataFrame(index=row_order, columns=DF_COLS, dtype="Float64")

    # ─── preload tenor values ──────────────────────────────────────────
    if preload_data:
        for row, payload in preload_data.items():
            if row not in row_order:
                continue
            if isinstance(payload, list) and len(payload) == 7:
                df.loc[row, TENOR_COLS] = payload
            elif isinstance(payload, dict):
                for tenor, val in payload.items():
                    if tenor in TENOR_COLS:
                        df.at[row, tenor] = val

    # ─── subtotal & total formulas ─────────────────────────────────────
    def recalc(d: pd.DataFrame) -> None:
        d["Total"] = d[TENOR_COLS].sum(axis=1, min_count=1, skipna=True)
        lcr = d[["O/N", "1D-1W", "1-2W", "2W-1M"]].sum(axis=1, min_count=1, skipna=True)
        d["LCR CF 30D"] = lcr
        d["W.LCR CF"]   = lcr

        start = 0
        for i, r in enumerate(row_order):
            if r in subtotal_rows:
                blk = row_order[start:i]
                d.loc[r, TENOR_COLS] = d.loc[blk, TENOR_COLS].sum()
                d.at[r, "Total"] = d.loc[r, TENOR_COLS].sum(skipna=True)
                l = d.loc[r, ["O/N", "1D-1W", "1-2W", "2W-1M"]].sum(skipna=True)
                d.at[r, "LCR CF 30D"] = l
                d.at[r, "W.LCR CF"]   = l
                start = i + 1

        d.loc[grand_total_row, TENOR_COLS] = d.loc[subtotal_rows, TENOR_COLS].sum()
        d.at[grand_total_row, "Total"] = d.loc[grand_total_row, TENOR_COLS].sum(skipna=True)
        g = d.loc[grand_total_row, ["O/N", "1D-1W", "1-2W", "2W-1M"]].sum(skipna=True)
        d.at[grand_total_row, "LCR CF 30D"] = g
        d.at[grand_total_row, "W.LCR CF"]   = g

    recalc(df)

    # ─── AG-Grid styling JS ────────────────────────────────────────────
    row_js = JsCode(f"""
function(p){{
  const greens={json.dumps(subtotal_rows)},
        blue=['{grand_total_row}'],
        m=p.data.Metric;
  if (blue.includes(m))  return {{'background-color':'#cce5ff','font-weight':'bold'}};
  if (greens.includes(m))return {{'background-color':'#d4f7d4','font-weight':'bold'}};
  const v=p.data.Total;
  if (v===null||v===undefined||v==='') return {{}};
  return {{'background-color': v<0 ? '#ffcccc' : '#d4f7d4'}};
}}""")

    pink_js = JsCode("function(){return {'background-color':'#ffe6f2'};}")
    grey_js = JsCode("function(){return {'background-color':'#f5f5f5'};}")

    # ─── GridOptions ───────────────────────────────────────────────────
    df_disp = df.reset_index(names="Metric")
    gb = GridOptionsBuilder.from_dataframe(df_disp)
    gb.configure_default_column(
        resizable=True,
        wrapText=True,
        autoHeight=True,
        flex=1,
        editable=_editable_js(["Total", *LCR_COLS], editable_row),
        sortable=False,
    )
    gb.configure_column("Metric", editable=False, flex=2)
    gb.configure_column("Total",  editable=False)
    gb.configure_column("CF",     cellStyle=grey_js)
    for col in LCR_COLS:
        gb.configure_column(col, editable=False, cellStyle=pink_js)
    gb.configure_grid_options(
        getRowStyle=row_js,
        domLayout="autoHeight",
        suppressMovableColumns=True,
    )

    # ─── render grid ───────────────────────────────────────────────────
    st.markdown(f"### {title}")
    grid = AgGrid(
        df_disp,
        gridOptions=gb.build(),
        fit_columns_on_grid_load=True,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        allow_unsafe_jscode=True,
    )

    df_updated = (
        pd.DataFrame(grid["data"])
        .set_index("Metric")
        .reindex(index=row_order, columns=DF_COLS)
        .astype("Float64")
    )
    recalc(df_updated)
    st.session_state[title] = df_updated

    # ─── Excel download button ─────────────────────────────────────────
    excel_bytes = _write_excel(
        df_updated,
        f"{title}.xlsx",
        subtotal_rows=subtotal_rows,
        grand_total_row=grand_total_row,
    )
    st.download_button(
        "⬇️ Download as Excel",
        data=excel_bytes,
        file_name=f"{title.replace(' ', '_')}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml"
            ".sheet"
        ),
    )


# ─────────────────────────────── demo tables ──────────────────────────
LENDING_ROWS = ROW_ORDER
SUBTOTALS    = [
    "Total Lending book", "Total MTM value", "Secured Lending",
    "Interbank Lending", "Intragroup Lending", "Total other assets",
]

build_subtotal_table(
    title="💰 Lending Book",
    row_order=LENDING_ROWS,
    subtotal_rows=SUBTOTALS,
    grand_total_row="Total Assets",
    preload_data={
        "Others": [1, 2, 1.5, 3, 4, 5, 6],
        "Interbank Lending": {"O/N": 9.9, "1-3M": 7.7},
        "Intragroup Lending": [0, 0, 1, 1, 2, 2, 3],
    },
    editable_row="iVision Loans",
)

st.divider()

build_subtotal_table(
    title="📊 Mini Demo",
    row_order=["Row A", "Row B", "Sub1", "Row C", "Sub2", "Grand"],
    subtotal_rows=["Sub1", "Sub2"],
    grand_total_row="Grand",
    preload_data={"Row A": [1, 1, 1, 1, 1, 1, 1]},
    editable_row="Row A",
)
