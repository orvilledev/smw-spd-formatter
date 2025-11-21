import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
import re

# --- Streamlit Page Setup ---
st.set_page_config(
    page_title="📦 SMW SPD Box Contents Formatter (Multi-File)",
    page_icon="📦",
    layout="wide",
)
st.title("📦 SMW SPD Box Contents Formatter (Multi-File)")
st.caption(
    "Upload multiple Excel files and consolidate them into one formatted output."
)

# --- File Uploader ---
uploaded_files = st.file_uploader(
    "📁 Upload up to 20 Excel Sheets",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) > 20:
        st.error("❌ You can only upload up to 20 Excel files.")
        st.stop()

    consolidated_contents = []
    consolidated_dims = []
    box_offset = 0

    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, header=10, engine="openpyxl")
            df.columns = df.columns.astype(str).str.strip()
        except Exception as e:
            st.error(f"❌ Error reading file {uploaded_file.name}: {e}")
            continue

        required_columns = ["UPC", "Box X", "Sku Units"]
        missing_cols = [c for c in required_columns if c not in df.columns]
        if missing_cols:
            st.warning(
                f"⚠️ File {uploaded_file.name} missing: {', '.join(missing_cols)}"
            )
            continue

        # --- Box Contents ---
        df_clean = df[required_columns].dropna(subset=["UPC", "Sku Units"])
        df_clean["UPC"] = (
            df_clean["UPC"]
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.replace("+", "", regex=False)
            .str.zfill(12)
        )
        df_clean["Sku Units"] = (
            pd.to_numeric(df_clean["Sku Units"], errors="coerce").fillna(0).astype(int)
        )
        df_clean.rename(
            columns={"Box X": "Box Number", "Sku Units": "Qty"}, inplace=True
        )

        # Sequential Box Numbering
        if len(df_clean) > 0:
            df_clean["Box Number"] = df_clean["Box Number"].astype(int) + box_offset
            max_box = df_clean["Box Number"].max()
            box_offset = max_box

        # Add Routing # from filename
        routing_number = uploaded_file.name.rsplit(".", 1)[0]
        df_clean["Routing #"] = routing_number

        consolidated_contents.append(df_clean)

        # --- Extract Dimensions with Routing # from filename ---
        dimension_pattern = r"\b\d{1,3}\.\d{1,2}X\d{1,3}\.\d{1,2}X\d{1,3}\.\d{1,2}\b"
        dimension_data = []
        for _, row in df.iterrows():
            for col in df.columns:
                val = str(row[col])
                if re.match(dimension_pattern, val):
                    length, width, height = val.split("X")
                    dimension_data.append((float(length), float(width), float(height)))
        if dimension_data:
            dim_df = pd.DataFrame(dimension_data, columns=["Length", "Width", "Height"])
            dim_df.insert(0, "Box Number", range(1, len(dim_df) + 1))
            dim_df.insert(4, "Routing #", routing_number)
            consolidated_dims.append(dim_df)

    # --- Final Assembly ---
    final_contents = pd.concat(consolidated_contents, ignore_index=True)
    final_dims = (
        pd.concat(consolidated_dims, ignore_index=True)
        if consolidated_dims
        else pd.DataFrame()
    )

    # --- Reset Box Number to start at 1 in All Box Dimensions ---
    if not final_dims.empty:
        final_dims["Box Number"] = range(1, len(final_dims) + 1)

    # --- Create Pivot Table (unique UPCs, summed quantities) ---
    all_grouped = final_contents.groupby(["UPC", "Box Number"], as_index=False)[
        "Qty"
    ].sum()
    final_pivot = pd.pivot_table(
        all_grouped,
        index="UPC",
        columns="Box Number",
        values="Qty",
        aggfunc="sum",
        fill_value=0,
    )
    if 0 in final_pivot.columns:
        final_pivot = final_pivot.drop(columns=0)
    final_pivot = final_pivot.reindex(sorted(final_pivot.columns), axis=1)
    final_pivot_display = final_pivot.replace(0, "")

    # --- Compute totals safely ---
    pivot_for_sum = final_pivot.fillna(0).astype(int)
    pivot_column_totals = pivot_for_sum.sum(axis=0)
    grand_total_value = pivot_column_totals.sum()

    # --- Save to Excel ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_contents.to_excel(writer, sheet_name="All Box Contents", index=False)
        if not final_dims.empty:
            final_dims.to_excel(writer, sheet_name="All Box Dimensions", index=False)
    output.seek(0)
    wb = load_workbook(output)

    # --- Styles ---
    header_fill = PatternFill(
        start_color="F0D1B4", end_color="F0D1B4", fill_type="solid"
    )
    routing_fill = PatternFill(
        start_color="EB89F0", end_color="EB89F0", fill_type="solid"
    )
    grand_total_value_fill = PatternFill(
        start_color="C2EDA6", end_color="C2EDA6", fill_type="solid"
    )
    bold_font = Font(bold=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    center_align = Alignment(horizontal="center", vertical="center")

    # --- All Box Contents ---
    ws = wb["All Box Contents"]

    # Format headers
    for cell in ws[1]:
        cell.font = bold_font
        cell.border = thin_border
        cell.alignment = center_align
        if cell.value == "Routing #":
            cell.fill = routing_fill
        else:
            cell.fill = header_fill

    # Pivot Table starting at G1
    start_col = 7
    ws.cell(row=1, column=start_col, value="UPC").font = bold_font
    ws.cell(row=1, column=start_col).alignment = center_align
    ws.cell(row=1, column=start_col).fill = header_fill
    ws.cell(row=1, column=start_col).border = thin_border

    for idx, col in enumerate(final_pivot_display.columns, start=start_col + 1):
        ws.cell(row=1, column=idx, value=f"Box {col}").font = bold_font
        ws.cell(row=1, column=idx).alignment = center_align
        ws.cell(row=1, column=idx).fill = header_fill
        ws.cell(row=1, column=idx).border = thin_border

    # Fill pivot table
    for r_idx, (upc, row) in enumerate(final_pivot_display.iterrows(), start=2):
        ws.cell(row=r_idx, column=start_col, value=upc).alignment = center_align
        for c_idx, value in enumerate(row.values, start=start_col + 1):
            ws.cell(row=r_idx, column=c_idx, value=value).alignment = center_align

    # Pivot totals & Grand Total
    pivot_total_row = len(final_pivot_display) + 2
    ws.cell(row=pivot_total_row, column=start_col, value="Total").font = bold_font
    ws.cell(row=pivot_total_row, column=start_col).alignment = center_align

    for idx, col in enumerate(final_pivot_display.columns, start=start_col + 1):
        ws.cell(
            row=pivot_total_row, column=idx, value=pivot_column_totals[col]
        ).font = bold_font
        ws.cell(row=pivot_total_row, column=idx).alignment = center_align

    grand_total_col_idx = start_col + len(final_pivot_display.columns) + 1
    ws.cell(row=1, column=grand_total_col_idx, value="Grand Total").font = bold_font
    ws.cell(row=1, column=grand_total_col_idx).alignment = center_align
    ws.cell(row=1, column=grand_total_col_idx).fill = routing_fill
    ws.cell(row=1, column=grand_total_col_idx).border = thin_border

    ws.cell(
        row=pivot_total_row, column=grand_total_col_idx, value=grand_total_value
    ).font = bold_font
    ws.cell(row=pivot_total_row, column=grand_total_col_idx).alignment = center_align
    ws.cell(
        row=pivot_total_row, column=grand_total_col_idx
    ).fill = grand_total_value_fill

    # --- Qty total & total number of boxes ---
    last_data_row = len(final_contents) + 1
    qty_total = final_contents["Qty"].sum()
    unique_boxes = final_contents["Box Number"].nunique()

    # Total Boxes
    ws.cell(row=last_data_row + 1, column=2, value="Total Boxes").font = bold_font
    ws.cell(row=last_data_row + 1, column=2).alignment = center_align
    ws.cell(row=last_data_row + 1, column=3, value=unique_boxes).font = bold_font
    ws.cell(row=last_data_row + 1, column=3).alignment = center_align
    ws.cell(row=last_data_row + 1, column=3).fill = grand_total_value_fill

    # Total Qty
    ws.cell(row=last_data_row + 2, column=2, value="Total Qty").font = bold_font
    ws.cell(row=last_data_row + 2, column=2).alignment = center_align
    ws.cell(row=last_data_row + 2, column=3, value=qty_total).font = bold_font
    ws.cell(row=last_data_row + 2, column=3).alignment = center_align
    ws.cell(row=last_data_row + 2, column=3).fill = grand_total_value_fill

    # Center all cells
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center_align

    # --- All Box Dimensions ---
    if not final_dims.empty:
        ws_dims = wb["All Box Dimensions"]
        for cell in ws_dims[1]:
            cell.font = bold_font
            cell.border = thin_border
            cell.alignment = center_align
            if cell.value == "Routing #":
                cell.fill = routing_fill
            else:
                cell.fill = header_fill
        for row in ws_dims.iter_rows():
            for cell in row:
                cell.alignment = center_align

    # --- Auto-adjust column widths dynamically ---
    for ws_iter in [ws] + ([wb["All Box Dimensions"]] if not final_dims.empty else []):
        for col in ws_iter.columns:
            max_len = max(
                len(str(cell.value)) if cell.value is not None else 0 for cell in col
            )
            ws_iter.column_dimensions[col[0].column_letter].width = max(max_len + 5, 12)

    # --- Save final Excel ---
    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    # --- Prepare combined filename with number of files ---
    num_files = len(uploaded_files)
    combined_filename = f"SMW-BC-Output-{num_files}-ITEMS.xlsx"

    # --- Download Button ---
    st.download_button(
        label="💾 Download Consolidated Formatted Output",
        data=final_output,
        file_name=combined_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
