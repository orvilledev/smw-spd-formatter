# app.py
import streamlit as st
import os
import zipfile
from io import BytesIO
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

# --- Page config (single call) ---
st.set_page_config(
    page_title="📦 Excel Tools & SMW Formatters",
    page_icon="📦",
    layout="wide",
)

# =========================================================
# SECTION 1: Excel Finder + ZIP Extractor (Cloud uploads only)
# =========================================================
st.header("🔎 EXCEL FINDER")
st.markdown(
    "Upload Excel / ZIP files and search filenames (or filenames inside ZIPs) "
    "using keywords. Matches can be downloaded as a combined ZIP."
)

uploaded_files_finder = st.file_uploader(
    "Upload Excel or ZIP files (for searching) — top section",
    type=["xlsx", "xls", "xlsm", "zip"],
    accept_multiple_files=True,
    key="finder_uploader",
)

patterns_input = st.text_area(
    "Enter filename keywords (one per line):",
    placeholder="Example:\nsales\nPO123\nreport",
    key="finder_patterns",
)

search_btn = st.button("🔍 Search Uploaded Files", key="search_btn")
excel_ext = (".xlsx", ".xls", ".xlsm")

if search_btn:
    patterns = [p.strip().lower() for p in patterns_input.splitlines() if p.strip()]
    if not patterns:
        st.error("⚠️ Please enter at least one keyword to search.")
    else:
        found_files = []
        if not uploaded_files_finder:
            st.warning("⚠️ No files uploaded for search.")
        else:
            for f in uploaded_files_finder:
                fname = f.name
                try:
                    fbytes = f.read()
                except Exception as e:
                    st.warning(f"⚠️ Could not read file {fname}: {e}")
                    continue

                fname_lower = fname.lower()
                # Direct Excel file match
                if fname_lower.endswith(excel_ext) and any(
                    p in fname_lower for p in patterns
                ):
                    found_files.append((fname, fbytes))
                # ZIP: inspect inside
                elif fname_lower.endswith(".zip"):
                    try:
                        with zipfile.ZipFile(BytesIO(fbytes)) as z:
                            for zi in z.namelist():
                                zi_lower = zi.lower()
                                if zi_lower.endswith(excel_ext) and any(
                                    p in zi_lower for p in patterns
                                ):
                                    extracted = z.read(zi)
                                    arcname = (
                                        f"{fname.split('.')[0]}_{zi.split('/')[-1]}"
                                    )
                                    found_files.append((arcname, extracted))
                    except zipfile.BadZipFile:
                        st.warning(f"⚠️ Cannot read ZIP file: {fname}")

        st.subheader("🔍 Search Results")
        if not found_files:
            st.error("❌ No matching Excel files found in the uploaded files.")
        else:
            for ff in found_files:
                st.write(f"🗂️ {ff[0]}")

            # prepare zip for download
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_out:
                for arcname, data in found_files:
                    zip_out.writestr(arcname, data)
            zip_buffer.seek(0)
            st.download_button(
                "📦 Download ZIP of Matches",
                data=zip_buffer,
                file_name="excel_search_results.zip",
                mime="application/zip",
            )

st.markdown("---")

# =========================================================
# SECTION 2: SMW SPD Box Contents Formatter (Option A logic)
# =========================================================
st.header("📑 SPD FORMATTER (SMW)")

st.markdown(
    "Upload up to 100 Excel or ZIP files (these may be the same files used in the top section). "
    "The app will consolidate All Box Contents and All Box Dimensions, add a pivot, styling, and a Summary sheet."
)

uploaded_files_formatter = st.file_uploader(
    "📁 Upload up to 100 Excel or ZIP files (for formatting) — middle section",
    type=["xlsx", "xls", "zip"],
    accept_multiple_files=True,
    key="formatter_uploader",
)

if uploaded_files_formatter:
    if len(uploaded_files_formatter) > 100:
        st.error("❌ You can only upload up to 100 files.")
        st.stop()

    all_files_to_process = []

    # extract excel files from uploaded files
    for uploaded_file in uploaded_files_formatter:
        file_name = uploaded_file.name
        try:
            file_bytes = uploaded_file.read()
        except Exception as e:
            st.warning(f"⚠️ Cannot read {file_name}: {e}")
            continue

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
    else:
        with st.spinner("Processing files..."):
            # processing variables
            consolidated_contents = []
            consolidated_dims = []
            box_offset = 0

            for fname, fbytes in all_files_to_process:
                try:
                    fbytes.seek(0)
                except:
                    pass

                try:
                    df = pd.read_excel(fbytes, header=10, engine="openpyxl")
                    df.columns = df.columns.astype(str).str.strip()
                except Exception as e:
                    st.warning(f"⚠️ Error reading file {fname}: {e}")
                    continue

                required_columns = ["UPC", "Box X", "Sku Units"]
                missing_cols = [c for c in required_columns if c not in df.columns]
                if missing_cols:
                    st.warning(f"⚠️ File {fname} missing: {', '.join(missing_cols)}")
                    continue

                # extract customer PO and routing # from Page1_1 or first sheet
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
                except Exception:
                    customer_po, routing_number = "", ""

                # --- Box Contents processing ---
                df_clean = (
                    df[required_columns].dropna(subset=["UPC", "Sku Units"]).copy()
                )
                df_clean["UPC"] = (
                    df_clean["UPC"]
                    .astype(str)
                    .str.replace(r"\.0$", "", regex=True)
                    .str.zfill(12)
                )
                df_clean["Sku Units"] = (
                    pd.to_numeric(df_clean["Sku Units"], errors="coerce")
                    .fillna(0)
                    .astype(int)
                )
                df_clean.rename(
                    columns={"Box X": "Box Number", "Sku Units": "Qty"}, inplace=True
                )

                if len(df_clean) > 0:
                    # offset Box Number so boxes from different files don't collide
                    df_clean["Box Number"] = (
                        df_clean["Box Number"].astype(int) + box_offset
                    )
                    box_offset = df_clean["Box Number"].max()

                df_clean["Customer PO"] = customer_po
                df_clean["Routing #"] = routing_number
                consolidated_contents.append(df_clean)

                # --- Dimensions extraction ---
                dimension_pattern = (
                    r"\b\d{1,3}\.\d{1,2}X\d{1,3}\.\d{1,2}X\d{1,3}\.\d{1,2}\b"
                )
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

                # --- Weight extraction (attempt to get bold cells in column G) ---
                try:
                    fbytes.seek(0)
                    wb_temp2 = load_workbook(fbytes, data_only=True)
                    ws_for_bold = (
                        wb_temp2["Page1_1"]
                        if "Page1_1" in wb_temp2.sheetnames
                        else wb_temp2[wb_temp2.sheetnames[0]]
                    )
                    bold_cells = [
                        cell
                        for cell in ws_for_bold["G"]
                        if getattr(cell.font, "bold", False) and cell.value is not None
                    ]
                    if bold_cells:
                        bold_cells = bold_cells[:-1]
                    weights = [cell.value for cell in bold_cells]
                except Exception:
                    weights = []

                num_boxes = max(len(weights), len(dimension_data), len(df_clean))
                start_box_num = (
                    box_offset - len(df_clean) + 1
                    if len(df_clean) > 0
                    else box_offset + 1
                )
                boxes = list(range(start_box_num, start_box_num + num_boxes))

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

            # --- Final assembly ---
            try:
                final_contents = (
                    pd.concat(consolidated_contents, ignore_index=True)
                    if consolidated_contents
                    else pd.DataFrame()
                )
            except ValueError:
                final_contents = pd.DataFrame()

            col_order = ["UPC", "Box Number", "Qty", "Customer PO", "Routing #"]
            if not final_contents.empty:
                final_contents = final_contents[
                    [c for c in col_order if c in final_contents.columns]
                    + [c for c in final_contents.columns if c not in col_order]
                ]
                final_contents = final_contents.sort_values(
                    by="Customer PO", ascending=True
                ).reset_index(drop=True)

                # Reassign Box Number based on Routing # groups
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

                # reorder to ensure Box Number is second column
                cols = list(final_contents.columns)
                if "Box Number" in cols:
                    cols.insert(1, cols.pop(cols.index("Box Number")))
                    final_contents = final_contents[cols]

                summary_df = final_contents[
                    ["Customer PO", "Routing #"]
                ].drop_duplicates(ignore_index=True)
            else:
                summary_df = pd.DataFrame()
                final_contents = pd.DataFrame()

            final_dims = (
                pd.concat(consolidated_dims, ignore_index=True)
                if consolidated_dims
                else pd.DataFrame()
            )
            if not final_dims.empty:
                final_dims = final_dims.dropna(
                    subset=["Weight", "Length", "Width", "Height"], how="any"
                ).reset_index(drop=True)
                final_dims = final_dims.drop_duplicates(
                    subset=["Routing #"], keep="first"
                ).reset_index(drop=True)

                summary_routing_order = (
                    summary_df["Routing #"].tolist() if not summary_df.empty else []
                )
                if summary_routing_order:
                    final_dims["Routing #"] = pd.Categorical(
                        final_dims["Routing #"],
                        categories=summary_routing_order,
                        ordered=True,
                    )
                    final_dims = final_dims.sort_values("Routing #").reset_index(
                        drop=True
                    )
                final_dims["Box Number"] = range(1, len(final_dims) + 1)

            # --- Write to Excel and style ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                if not final_contents.empty:
                    final_contents.to_excel(
                        writer, sheet_name="All Box Contents", index=False
                    )
                else:
                    pd.DataFrame(
                        columns=["UPC", "Box Number", "Qty", "Customer PO", "Routing #"]
                    ).to_excel(writer, sheet_name="All Box Contents", index=False)

                if not final_dims.empty:
                    final_dims.to_excel(
                        writer, sheet_name="All Box Dimensions", index=False
                    )

            output.seek(0)
            wb = load_workbook(output)

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

            # Style All Box Contents headers
            ws = wb["All Box Contents"]
            for cell in ws[1]:
                cell.font = bold_font
                cell.border = thin_border
                cell.fill = special_fill if (cell.value == "Routing #") else header_fill

            # create pivot if data exists
            if not final_contents.empty:
                all_grouped = final_contents.groupby(
                    ["UPC", "Box Number"], as_index=False
                )["Qty"].sum()
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
                for idx, col in enumerate(
                    final_pivot_display.columns, start=start_col + 1
                ):
                    ws.cell(row=1, column=idx, value=f"Box {col}").font = bold_font
                    ws.cell(row=1, column=idx).fill = header_fill
                    ws.cell(row=1, column=idx).border = thin_border

                for r_idx, (upc, row) in enumerate(
                    final_pivot_display.iterrows(), start=2
                ):
                    ws.cell(row=r_idx, column=start_col, value=upc)
                    for c_idx, value in enumerate(row.values, start=start_col + 1):
                        ws.cell(row=r_idx, column=c_idx, value=value)

                total_col_idx = start_col + len(final_pivot_display.columns) + 1
                ws.cell(
                    row=1, column=total_col_idx, value="Total per UPC"
                ).font = bold_font
                ws.cell(row=1, column=total_col_idx).fill = special_fill

                for r_idx, total in enumerate(pivot_row_totals, start=2):
                    ws.cell(row=r_idx, column=total_col_idx, value=total)

                pivot_total_row = len(final_pivot_display) + 2
                ws.cell(
                    row=pivot_total_row, column=start_col, value="Total per Box"
                ).font = bold_font
                ws.cell(row=pivot_total_row, column=start_col).fill = special_fill

                for idx, col in enumerate(
                    final_pivot_display.columns, start=start_col + 1
                ):
                    ws.cell(
                        row=pivot_total_row, column=idx, value=pivot_column_totals[col]
                    )

                ws.cell(
                    row=pivot_total_row, column=total_col_idx, value=grand_total_value
                ).font = bold_font
                ws.cell(row=pivot_total_row, column=total_col_idx).fill = special_fill

            # Check if Customer PO column D is alphabetically sorted
            ws_contents = wb["All Box Contents"]
            customer_po_values = []
            for row in ws_contents.iter_rows(min_row=2, min_col=4, max_col=4):
                value = row[0].value
                if value not in [None, ""]:
                    customer_po_values.append(str(value))

            last_row = ws_contents.max_row

            if customer_po_values == sorted(
                customer_po_values, key=lambda x: x.lower()
            ):
                status_text = "✔ POs are alphabetically arranged"
                status_color = "92d050"
            else:
                status_text = "❌ POs are NOT alphabetically arranged"
                status_color = "ff0000"

            status_cell = ws_contents.cell(
                row=last_row + 2, column=4, value=status_text
            )
            status_cell.font = Font(bold=True, color="000000")
            status_cell.fill = PatternFill(
                start_color=status_color, end_color=status_color, fill_type="solid"
            )
            status_cell.border = thin_border
            status_cell.alignment = center_align

            # Style All Box Dimensions if present
            if "All Box Dimensions" in wb.sheetnames and not final_dims.empty:
                ws_dims = wb["All Box Dimensions"]
                for cell in ws_dims[1]:
                    cell.font = bold_font
                    cell.border = thin_border
                    cell.fill = (
                        special_fill if (cell.value == "Routing #") else header_fill
                    )

            # Create summary sheet
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

            # Center align and set column widths
            for ws_iter in wb.worksheets:
                for row in ws_iter.iter_rows():
                    for cell in row:
                        cell.alignment = center_align
                        if cell.border is None:
                            cell.border = thin_border
                for col in ws_iter.columns:
                    try:
                        max_len = max(
                            len(str(cell.value)) if cell.value is not None else 0
                            for cell in col
                        )
                    except Exception:
                        max_len = 0
                    try:
                        ws_iter.column_dimensions[col[0].column_letter].width = max(
                            max_len + 5, 12
                        )
                    except Exception:
                        pass

            # Save and provide download
            final_output = BytesIO()
            wb.save(final_output)
            final_output.seek(0)

            combined_filename = f"SMW-BC-Output-{len(all_files_to_process)}-ITEMS.xlsx"
            st.download_button(
                label="💾 Download Consolidated Formatted Output (Option A)",
                data=final_output,
                file_name=combined_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.success("✅ Processing complete. Download the formatted workbook above.")

st.markdown("---")

# =========================================================
# SECTION 3: Single-file SMW Box Contents Formatter (Code 3)
# =========================================================
st.header("📑 LTL FORMATTER (SMW)")

st.markdown(
    "Upload a single Excel file. This section processes one file, creates Box Contents, Pivot, and Box Dimensions, and shows metrics + previews."
)

uploaded_file_single = st.file_uploader(
    "📁 Select an Excel file (single-file formatter) — bottom section",
    type=["xlsx", "xls"],
    key="single_uploader",
)

if uploaded_file_single:
    try:
        input_filename = uploaded_file_single.name
        base_name, ext = os.path.splitext(input_filename)
        output_filename = f"{base_name} formatted{ext}"

        # Read with header row 11 (header=10)
        df = pd.read_excel(uploaded_file_single, header=10, engine="openpyxl")
        df.columns = df.columns.astype(str).str.strip()
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {e}")
    else:
        required_columns = ["UPC", "Box X", "Sku Units"]
        missing_cols = [c for c in required_columns if c not in df.columns]

        if missing_cols:
            st.warning(f"⚠️ Missing columns: {', '.join(missing_cols)}")
        else:
            # Box Contents
            df_clean = df[required_columns].dropna(subset=["UPC", "Sku Units"]).copy()
            df_clean["UPC"] = (
                df_clean["UPC"]
                .astype(str)
                .str.replace(r"\.0$", "", regex=True)
                .str.replace("+", "", regex=False)
                .str.zfill(12)
            )
            df_clean["Sku Units"] = (
                pd.to_numeric(df_clean["Sku Units"], errors="coerce")
                .fillna(0)
                .astype(int)
            )
            df_clean.rename(
                columns={"Box X": "Box Number", "Sku Units": "Qty"}, inplace=True
            )

            # Pivot Table
            pivot_table = pd.pivot_table(
                df_clean,
                index="UPC",
                columns="Box Number",
                values="Qty",
                aggfunc="sum",
                fill_value=0,
            ).reset_index()
            pivot_table = pivot_table.replace(0, "")

            # Totals
            total_qty = int(df_clean["Qty"].sum()) if not df_clean.empty else 0
            total_boxes = (
                int(df_clean["Box Number"].nunique()) if not df_clean.empty else 0
            )

            # Extract Carton Weights
            carton_weights = []
            try:
                # load workbook directly from uploaded file-like
                uploaded_file_single.seek(0)
                wb_input = load_workbook(uploaded_file_single, data_only=True)
                ws_page1 = (
                    wb_input["Page1_1"]
                    if "Page1_1" in wb_input.sheetnames
                    else wb_input[wb_input.sheetnames[0]]
                )
                for row in ws_page1.iter_rows(min_row=1, max_col=7):
                    cell = row[6]  # column G (index 6)
                    if getattr(cell.font, "bold", False) and isinstance(
                        cell.value, (int, float)
                    ):
                        carton_weights.append(cell.value)
                if carton_weights:
                    carton_weights = carton_weights[:-1]
            except Exception:
                carton_weights = []

            total_carton_weight = sum(
                [w for w in carton_weights if isinstance(w, (int, float))]
            )
            total_carton_weight_plus35 = total_carton_weight + 35

            # Extract Dimensions
            dimension_pattern = (
                r"\b\d{1,3}\.\d{1,2}X\d{1,3}\.\d{1,2}X\d{1,3}\.\d{1,2}\b"
            )
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

            # Box Dimensions DataFrame
            dim_df = pd.DataFrame()
            if dimension_data:
                dim_df = pd.DataFrame(
                    dimension_data, columns=["Length", "Width", "Height"]
                )
                dim_df.insert(0, "Box Number", range(1, len(dim_df) + 1))
                weights_column = carton_weights[: len(dim_df)] + [""] * max(
                    0, len(dim_df) - len(carton_weights)
                )
                dim_df.insert(1, "Carton Weight", weights_column)

            # Write to Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_clean.to_excel(writer, sheet_name="Box Contents", index=False)
                pivot_table.to_excel(writer, sheet_name="Pivot Table", index=False)
                if not dim_df.empty:
                    dim_df.to_excel(writer, sheet_name="Box Dimensions", index=False)

            output.seek(0)
            wb = load_workbook(output)

            yellow_fill = PatternFill(
                start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
            )
            header_font = Font(bold=True, size=14)
            align_center = Alignment(horizontal="center", vertical="center")
            thin_border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )

            def style_sheet(ws, keep_decimals=False, force_int_cols=[]):
                # header row style
                for row in ws.iter_rows(min_row=1, max_row=1):
                    for cell in row:
                        cell.fill = yellow_fill
                        cell.font = header_font
                        cell.alignment = align_center
                        cell.border = thin_border
                # body rows
                for row in ws.iter_rows(
                    min_row=2, max_row=ws.max_row, max_col=ws.max_column
                ):
                    for col_idx, cell in enumerate(row, start=1):
                        cell.border = thin_border
                        cell.alignment = align_center
                        if isinstance(cell.value, (int, float)):
                            if col_idx in force_int_cols:
                                cell.number_format = "0"
                            elif keep_decimals:
                                cell.number_format = "0.00"
                            else:
                                cell.number_format = "0"
                # column widths
                for col_idx in range(1, ws.max_column + 1):
                    ws.column_dimensions[get_column_letter(col_idx)].width = 18

            # Apply formatting
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                if sheet_name == "Box Dimensions":
                    style_sheet(ws, keep_decimals=True, force_int_cols=[1])
                elif sheet_name == "Pivot Table":
                    style_sheet(ws, keep_decimals=False)
                    # Reduce width to 1/4 excluding UPC
                    for col_idx in range(2, ws.max_column + 1):
                        try:
                            current_width = ws.column_dimensions[
                                get_column_letter(col_idx)
                            ].width
                            ws.column_dimensions[
                                get_column_letter(col_idx)
                            ].width = max(4, (current_width or 18) / 4)
                        except Exception:
                            ws.column_dimensions[get_column_letter(col_idx)].width = 4
                    ws.column_dimensions["A"].width = 25
                else:
                    style_sheet(ws, keep_decimals=False)

            # Totals for Box Contents
            ws_contents = wb["Box Contents"]
            total_row = ws_contents.max_row + 2
            ws_contents[f"A{total_row}"] = "Total Qty:"
            ws_contents[f"B{total_row}"] = total_qty
            ws_contents[f"A{total_row + 1}"] = "Total Boxes:"
            ws_contents[f"B{total_row + 1}"] = total_boxes
            for r in range(total_row, total_row + 2):
                for c in range(1, 3):
                    cell = ws_contents.cell(row=r, column=c)
                    cell.font = Font(bold=True)
                    cell.border = thin_border
                    cell.alignment = align_center

            # Total Carton Weight in Box Dimensions
            if "Box Dimensions" in wb.sheetnames:
                ws_dim = wb["Box Dimensions"]
                carton_col = 2
                last_row = ws_dim.max_row
                total_weight = sum(
                    [
                        ws_dim.cell(row=r, column=carton_col).value
                        for r in range(2, last_row + 1)
                        if isinstance(
                            ws_dim.cell(row=r, column=carton_col).value, (int, float)
                        )
                    ]
                )
                total_weight += 35
                total_row = last_row + 1
                ws_dim.cell(row=total_row, column=1, value="Total Carton Weight (+35):")
                ws_dim.cell(row=total_row, column=carton_col, value=total_weight)
                for col in [1, carton_col]:
                    cell = ws_dim.cell(row=total_row, column=col)
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = thin_border
                ws_dim.column_dimensions["A"].width = 30

            formatted_output = BytesIO()
            wb.save(formatted_output)
            formatted_output.seek(0)

            # Streamlit Download
            st.download_button(
                label=f"💾 Download {output_filename}",
                data=formatted_output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
