import streamlit as st
import pandas as pd
import tempfile
import os
import re
from pathlib import Path

# Import the merge utility functions
import merge

st.title("CAPA Log Merger")

uploaded_files = st.file_uploader(
    "Drag & drop your CAPA Excel files (CAR, PTO, PAR, CAF) here",
    type=["xlsx"],
    accept_multiple_files=True,
)

if uploaded_files:
    results = {"CAR": [], "PTO": [], "PAR": [], "CAF": []}
    warnings = []

    for uploaded in uploaded_files:
        # Save to a temporary file so we can use merge.read_sheet
        temp_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp_path.write(uploaded.getvalue())
        temp_path.close()
        file_name = uploaded.name
        st.write(f"Processing **{file_name}**…")

        try:
            xl = pd.ExcelFile(temp_path.name)
        except Exception as e:
            warnings.append(f"⚠  Could not open {file_name}: {e}")
            os.unlink(temp_path.name)
            continue

        # Try to extract year from file name (if any)
        yr_match = re.search(r"(20\d{2})", file_name)
        year = int(yr_match.group(1)) if yr_match else 0

        for doc_type in ["CAR", "PTO", "PAR", "CAF"]:
            sheet_name = merge.find_sheet(list(xl.sheet_names), doc_type)
            if not sheet_name:
                warnings.append(f"⚠  No {doc_type} sheet found in {file_name}")
                continue
            df, errs = merge.read_sheet(temp_path.name, sheet_name, doc_type, year)
            warnings.extend(errs)
            if df is not None and not df.empty:
                results[doc_type].append(df)
                closed = (df["status"] == "CLOSED").sum()
                open_ = (df["status"] == "OPEN").sum()
                st.write(f"✓ {doc_type:4s} – {len(df):4d} rows  ({closed} closed, {open_} open)  [{sheet_name}]")
            else:
                warnings.append(f"⚠  {doc_type} in {file_name} returned 0 usable rows")
        os.unlink(temp_path.name)

    # Combine & deduplicate
    st.write("\nCombining and deduplicating…")
    combined = {}
    num_col = {"CAR": "car_number", "PTO": "pto_number", "PAR": "par_number", "CAF": "caf_number"}

    for doc_type, frames in results.items():
        if not frames:
            st.warning(f"⚠  No data collected for {doc_type}")
            combined[doc_type] = pd.DataFrame()
            continue
        df = pd.concat(frames, ignore_index=True)
        id_col = num_col[doc_type]
        if id_col in df.columns and "location" in df.columns:
            before = len(df)
            df = (
                df.sort_values("year", ascending=False)
                .drop_duplicates(subset=[id_col, "location"], keep="first")
                .sort_values(["year", id_col], ascending=[True, True])
            )
            dupes = before - len(df)
            if dupes > 0:
                st.info(f"↳ {doc_type}: removed {dupes} duplicate record(s)")
        combined[doc_type] = df
        st.write(f"✓ {doc_type:4s} – {len(df):4d} total rows across all years")

    # Prepare output as an in-memory Excel file
    output_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    out_path = output_buffer.name
    output_buffer.close()

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for doc_type, df in combined.items():
            if df.empty:
                continue
            sheet = f"Data - {doc_type}s"
            df.to_excel(writer, sheet_name=sheet, index=False)
        # Summary sheet
        summary_rows = []
        for doc_type, df in combined.items():
            if df.empty:
                continue
            for yr in sorted(df["year"].unique()):
                ydf = df[df["year"] == yr]
                summary_rows.append(
                    {
                        "Doc Type": doc_type,
                        "Year": yr,
                        "Total": len(ydf),
                        "Closed": (ydf["status"] == "CLOSED").sum(),
                        "Open": (ydf["status"] == "OPEN").sum(),
                        "Avg Days to Close": round(
                            ydf.loc[ydf["status"] == "CLOSED", "days_to_close"].mean(),
                            1,
                        )
                        if (ydf["status"] == "CLOSED").any()
                        else "N/A",
                    }
                )
        if summary_rows:
            pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)

    # Read the output into memory for download
    with open(out_path, "rb") as f:
        download_data = f.read()
    os.unlink(out_path)

    st.success("Done! The merged file is ready for download.")

    # Display summary data
    if combined:
        for doc_type, df in combined.items():
            if not df.empty:
                st.subheader(f"{doc_type} – {len(df)} rows")
                st.dataframe(df.head(10))

    # Show warnings if any
    if warnings:
        st.error("Warnings:")
        for w in warnings:
            st.write(w)

    st.download_button(
        label="Download Merged CAPA Master (xlsx)",
        data=download_data,
        file_name="CAPA_Master.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.write("Upload one or more Excel files to merge them.")
