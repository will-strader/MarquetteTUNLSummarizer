import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime

def col_to_idx(col):
    import string
    col = col.upper().strip()
    idx = 0
    for c in col:
        if c in string.ascii_uppercase:
            idx = idx * 26 + (ord(c) - ord('A') + 1)
    return idx - 1

st.title("Rat Performance Analyzer")

uploaded_files = st.file_uploader("Upload CSV or Excel files", type=["csv", "xlsx"], accept_multiple_files=True)

with st.expander("Advanced Settings"):
    animal_col = st.text_input("Animal ID column", value="J")
    correct_col = st.text_input("NumCorrect column", value="AP")
    trial_col = st.text_input("Trial# column", value="AQ")
    dist_col = st.text_input("DistanceGP column", value="AR")
    range_input = st.text_input("Distance ranges (e.g., 1-4,5-8,9-13)", value="1-4,5-8,9-13")

if uploaded_files:
    a_idx = col_to_idx(animal_col)
    c_idx = col_to_idx(correct_col)
    t_idx = col_to_idx(trial_col)
    d_idx = col_to_idx(dist_col)
    ranges = []
    for part in range_input.split(','):
        try:
            mn, mx = map(float, part.strip().split('-'))
            ranges.append((mn, mx))
        except:
            st.error(f"Invalid range: {part}")

    records = {}

    for f in uploaded_files:
        ext = f.name.split('.')[-1].lower()
        if ext == 'csv':
            df = pd.read_csv(f, header=0, dtype=str)
        else:
            df = pd.read_excel(f, sheet_name=0, dtype=str)

        for idx, row in df.iterrows():
            try:
                aid = row.iat[a_idx]
                trial = row.iat[t_idx]
                if pd.isna(aid) or pd.isna(trial):
                    continue
                if aid not in records:
                    records[aid] = {}
                if trial in records[aid]:
                    continue
                dist = float(row.iat[d_idx])
                corr = int(float(row.iat[c_idx]))
                excel_row = idx + 2
                records[aid][trial] = {
                    'dist': dist,
                    'corr': corr,
                    'coord': (f"{correct_col.upper()}{excel_row}", f"{dist_col.upper()}{excel_row}")
                }
            except:
                continue

    out_rows = []
    for aid, trials in records.items():
        row = {'Animal ID': aid}
        for mn, mx in ranges:
            sel = [v for v in trials.values() if mn <= v['dist'] <= mx]
            count = len(sel)
            total = sum(v['corr'] for v in sel)
            pct = (total / count * 100) if count > 0 else 0.0
            coords = ";".join(f"[{v['coord'][0]},{v['coord'][1]}]" for v in sel)
            diag = f"{total},{count},{coords}"
            row[f"%C {mn:g}-{mx:g}"] = pct
            row[f"Diag {mn:g}-{mx:g}"] = diag
        out_rows.append(row)

    result_df = pd.DataFrame(out_rows).sort_values("Animal ID")
    st.dataframe(result_df)

    out_buffer = io.BytesIO()
    output_filename = f"rat_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    result_df.to_excel(out_buffer, index=False)
    st.download_button("Download Summary as Excel", data=out_buffer.getvalue(), file_name=output_filename)