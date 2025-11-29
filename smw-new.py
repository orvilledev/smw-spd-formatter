import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import re
import zipfile

# --- Streamlit Page Setup ---
st.set_page_config(
    page_title="📦 SMW SPD Box Contents Formatter (Multi-File + ZIP)",
    page_icon="📦",
    layout="wide",
)
st.title("📦 SMW SPD Box Contents Formatter (Multi-File + ZIP)")
st.caption(
    "Upload multiple Excel or ZIP files and consolidate them into one formatted output."
)

# --- File Uploader ---
uploaded_files = st.file_uploader(
    "📁 Upload up to 20 Excel or ZIP files",
    type=["xlsx", "xls", "zip"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) > 20:
        st.error("❌ You can only upload up to 20 files.")
        st.stop()

    all_files_to_process = []

    # --- Extract Excel files from uploaded files ---
    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        file_bytes = uploaded_file.read()
        uploaded_file.seek(0)

        if file_name.lower().endswith((".xlsx", ".xls")):
            all_files_to_process.append((file_name, BytesIO(file_bytes)))
        elif file_name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(BytesIO(file_bytes)) as z:
                    for zip_item in z.namelist():
                        if zip_item.lower().endswith((".xlsx", ".xls")):
                            extracted_bytes = z.read(zip_item)
                            arcname = (
                                f"{file_name.split('.')[0]}_{zip_item.split('/')[-1]}"
                            )
                            all_files_to_process.append(
                                (arcname, BytesIO(extracted_bytes))
                            )
            except zipfile.BadZipFile:
                st.warning(f"⚠️ Cannot read ZIP file: {file_name}")

    if len(all_files_to_process) == 0:
        st.warning("No valid Excel files found in the uploaded files.")
        st.stop()

    # --- Initialize containers ---
    consolidated_contents = []
    consolidated_dims = []
    box_offset = 0

    # --- Process each Excel file ---
    for fname, fbytes in all_files_to_process:
        try:
            df = pd.read_excel(fbytes, header=10, engine="openpyxl")
            df.columns = df.columns.astype(str).str.strip()
        except Exception as e:
            st.error(f"❌ Error reading file {fname}: {e}")
            continue

        required_columns = ["UPC", "Box X", "Sku Units"]
        missing_cols = [c for c in required_columns if c not in df.columns]
        if missing_cols:
            st.warning(f"⚠️ File {fname} missing: {', '.join(missing_cols)}")
            continue

        # --- Extract Customer PO and Routing # ---
        try:
            fbytes.seek(0)
            wb_temp = load_workbook(fbytes, data_only=True)
            ws_temp = (
                wb_temp["Page1_1"]
                if "Page1_1" in wb_temp.sheetnames
                else wb_temp[wb_temp.sheetnames[0]]
            )
            customer_po = ws_temp["G5"].value or ""
            routing_number = ws_temp["G6"].value or ""
        except:
            customer_po, routing_number = "", ""

        # --- Box Contents ---
        df_clean = df[required_columns].dropna(subset=["UPC", "Sku Units"]).copy()
        df_clean["UPC"] = (
            df_clean["UPC"]
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.zfill(12)
        )
        df_clean["Sku Units"] = (
            pd.to_numeric(df_clean["Sku Units"], errors="coerce").fillna(0).astype(int)
        )
        df_clean.rename(
            columns={"Box X": "Box Number", "Sku Units": "Qty"}, inplace=True
        )

        if len(df_clean) > 0:
            df_clean["Box Number"] = df_clean["Box Number"].astype(int) + box_offset
            box_offset = df_clean["Box Number"].max()

        df_clean["Customer PO"] = customer_po
        df_clean["Routing #"] = routing_number
        consolidated_contents.append(df_clean)

        # --- Extract Dimensions ---
        dimension_pattern = r"\b\d{1,3}\.\d{1,2}X\d{1,3}\.\d{1,2}X\d{1,3}\.\d{1,2}\b"
        dimension_data = []
        for _, row in df.iterrows():
            for col in df.columns:
                val = str(row[col])
                if re.match(dimension_pattern, val):
                    try:
                        length, width, height = val.split("X")
                        dimension_data.append(
                            (float(length), float(width), float(height))
                        )
                    except:
                        pass

        # --- Extract Weight ---
        try:
            bold_cells = [
                cell
                for cell in ws_temp["G"]
                if cell.font.bold and cell.value is not None
            ]
            if bold_cells:
                bold_cells = bold_cells[:-1]
            weights = [cell.value for cell in bold_cells]
        except:
            weights = []

        num_boxes = max(len(weights), len(dimension_data), len(df_clean))
        boxes = range(
            box_offset - len(df_clean) + 1, box_offset - len(df_clean) + 1 + num_boxes
        )

        lengths, widths, heights = (
            zip(*dimension_data) if dimension_data else ([], [], [])
        )
        lengths = list(lengths) + [""] * (num_boxes - len(lengths))
        widths = list(widths) + [""] * (num_boxes - len(widths))
        heights = list(heights) + [""] * (num_boxes - len(heights))
        weights = list(weights) + [""] * (num_boxes - len(weights))

        df_dims = pd.DataFrame(
            {
                "Box Number": boxes,
                "Weight": weights,
                "Length": lengths,
                "Width": widths,
                "Height": heights,
                "Routing #": routing_number,
            }
        )
        consolidated_dims.append(df_dims)

    # --- Final Assembly ---
    final_contents = pd.concat(consolidated_contents, ignore_index=True)
    col_order = ["UPC", "Box Number", "Qty", "Customer PO", "Routing #"]
    final_contents = final_contents[
        col_order + [c for c in final_contents.columns if c not in col_order]
    ]
    final_contents = final_contents.sort_values(
        by="Customer PO", ascending=True
    ).reset_index(drop=True)

    # --- Reassign Box Number based on Routing # groups ---
    routing_series = final_contents["Routing #"].fillna("").astype(str)
    routing_to_group = {}
    group_list = []
    group_counter = 1
    for r in routing_series:
        if r not in routing_to_group:
            routing_to_group[r] = group_counter
            group_counter += 1
        group_list.append(routing_to_group[r])
    final_contents["Box Number"] = group_list
    final_contents["Box Number"] = final_contents["Box Number"].astype(int)
    cols = list(final_contents.columns)
    if "Box Number" in cols:
        cols.insert(1, cols.pop(cols.index("Box Number")))
        final_contents = final_contents[cols]

    # --- Summary DataFrame ---
    summary_df = final_contents[["Customer PO", "Routing #"]].drop_duplicates(
        ignore_index=True
    )

    # --- Final Dimensions Processing ---
    final_dims = (
        pd.concat(consolidated_dims, ignore_index=True)
        if consolidated_dims
        else pd.DataFrame()
    )
    if not final_dims.empty:
        # 1. Remove rows with empty Weight, Length, Width, or Height
        final_dims = final_dims.dropna(
            subset=["Weight", "Length", "Width", "Height"], how="any"
        ).reset_index(drop=True)

        # 2. Remove duplicate Routing # (keep first occurrence)
        final_dims = final_dims.drop_duplicates(
            subset=["Routing #"], keep="first"
        ).reset_index(drop=True)

        # 3. Order Routing # to match Summary tab
        summary_routing_order = summary_df["Routing #"].tolist()
        final_dims["Routing #"] = pd.Categorical(
            final_dims["Routing #"], categories=summary_routing_order, ordered=True
        )
        final_dims = final_dims.sort_values("Routing #").reset_index(drop=True)

        # 4. Reassign sequential Box Numbers
        final_dims["Box Number"] = range(1, len(final_dims) + 1)

    # --- Write Excel ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_contents.to_excel(writer, sheet_name="All Box Contents", index=False)
        if not final_dims.empty:
            final_dims.to_excel(writer, sheet_name="All Box Dimensions", index=False)
    output.seek(0)
    wb = load_workbook(output)

    # --- Styles & Pivot Table & Summary ---
    header_fill = PatternFill(
        start_color="e3d3a8", end_color="e3d3a8", fill_type="solid"
    )
    special_fill = PatternFill(
        start_color="e09ddf", end_color="e09ddf", fill_type="solid"
    )
    bold_font = Font(bold=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    center_align = Alignment(horizontal="center", vertical="center")

    # --- Style All Box Contents headers ---
    ws = wb["All Box Contents"]
    for cell in ws[1]:
        cell.font = bold_font
        cell.border = thin_border
        cell.fill = special_fill if cell.value == "Routing #" else header_fill

    # --- Create Pivot Table ---
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
    final_pivot_display = final_pivot.replace(0, "")
    pivot_for_sum = final_pivot.fillna(0).astype(int)
    pivot_column_totals = pivot_for_sum.sum(axis=0)
    pivot_row_totals = pivot_for_sum.sum(axis=1)
    grand_total_value = pivot_row_totals.sum()

    start_col = 10  # Column J
    ws.cell(row=1, column=start_col, value="UPC").font = bold_font
    ws.cell(row=1, column=start_col).fill = header_fill

    for idx, col in enumerate(final_pivot_display.columns, start=start_col + 1):
        ws.cell(row=1, column=idx, value=f"Box {col}").font = bold_font
        ws.cell(row=1, column=idx).fill = header_fill
        ws.cell(row=1, column=idx).border = thin_border

    for r_idx, (upc, row) in enumerate(final_pivot_display.iterrows(), start=2):
        ws.cell(row=r_idx, column=start_col, value=upc)
        for c_idx, value in enumerate(row.values, start=start_col + 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    total_col_idx = start_col + len(final_pivot_display.columns) + 1
    ws.cell(row=1, column=total_col_idx, value="Total per UPC").font = bold_font
    ws.cell(row=1, column=total_col_idx).fill = special_fill

    for r_idx, total in enumerate(pivot_row_totals, start=2):
        ws.cell(row=r_idx, column=total_col_idx, value=total)

    pivot_total_row = len(final_pivot_display) + 2
    ws.cell(
        row=pivot_total_row, column=start_col, value="Total per Box"
    ).font = bold_font
    ws.cell(row=pivot_total_row, column=start_col).fill = special_fill

    for idx, col in enumerate(final_pivot_display.columns, start=start_col + 1):
        ws.cell(row=pivot_total_row, column=idx, value=pivot_column_totals[col])
    ws.cell(
        row=pivot_total_row, column=total_col_idx, value=grand_total_value
    ).font = bold_font
    ws.cell(row=pivot_total_row, column=total_col_idx).fill = special_fill

    # --- Style Dimensions ---
    if not final_dims.empty:
        ws_dims = wb["All Box Dimensions"]
        for cell in ws_dims[1]:
            cell.font = bold_font
            cell.border = thin_border
            cell.fill = special_fill if cell.value == "Routing #" else header_fill

    # --- Summary Sheet ---
    ws_summary = wb.create_sheet(title="Summary", index=2)
    for idx, col_name in enumerate(summary_df.columns, start=1):
        cell = ws_summary.cell(row=1, column=idx, value=col_name)
        cell.font = bold_font
        cell.border = thin_border
        cell.fill = special_fill if col_name == "Routing #" else header_fill

    for r_idx, row in enumerate(
        dataframe_to_rows(summary_df, index=False, header=False), start=2
    ):
        for c_idx, value in enumerate(row, start=1):
            ws_summary.cell(row=r_idx, column=c_idx, value=value)

    # --- Center align + Auto column widths ---
    for ws_iter in wb.worksheets:
        for row in ws_iter.iter_rows():
            for cell in row:
                cell.alignment = center_align
        for col in ws_iter.columns:
            max_len = max(
                len(str(cell.value)) if cell.value is not None else 0 for cell in col
            )
            ws_iter.column_dimensions[col[0].column_letter].width = max(max_len + 5, 12)

    # --- Save final Excel ---
    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    combined_filename = f"SMW-BC-Output-{len(all_files_to_process)}-ITEMS.xlsx"
    st.download_button(
        label="💾 Download Consolidated Formatted Output",
        data=final_output,
        file_name=combined_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
