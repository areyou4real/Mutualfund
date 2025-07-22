# -*- coding: utf-8 -*-
import pandas as pd
from io import BytesIO
import re

# ---------------------------- FUND PROCESSORS ----------------------------

def process_adityabirla(file_bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheet = xls.sheet_names[0]
    df = xls.parse(sheet)

    if df.shape[1] < 7:
        raise Exception("Not enough columns to extract 2nd and 7th.")

    df = df.iloc[:, [1, 6]]
    df.columns = df.columns.str.strip()
    df = df.dropna(how='all').reset_index(drop=True)

    while df.shape[1] < 4:
        df[f"extra_{df.shape[1] + 1}"] = None

    insert_values_col4 = []
    insert_labels_col3 = []

    # === Derivatives ===
    section_sum = 0
    count = 0
    for idx in range(len(df)):
        val = str(df.iloc[idx, 0]).lower()
        if "disclosure in derivatives" in val:
            j = idx
            while j < len(df):
                cell_val = df.iloc[j, 1]
                if pd.isna(cell_val) and count > 0:
                    break
                elif pd.notna(cell_val):
                    try:
                        section_sum += float(cell_val)
                        count += 1
                    except:
                        pass
                j += 1
            break
    if count > 0:
        insert_labels_col3.append("Hedged Equity")
        insert_values_col4.append(section_sum)

    # === Gold ===
    for idx, row in df.iterrows():
        desc = str(row.iloc[0]).lower()
        if "gold" in desc:
            try:
                insert_labels_col3.append("Gold")
                insert_values_col4.append(float(row.iloc[1]))
            except:
                pass
            break

    # === Silver ===
    for idx, row in df.iterrows():
        desc = str(row.iloc[0]).lower()
        if "silver" in desc:
            try:
                insert_labels_col3.append("Silver")
                insert_values_col4.append(float(row.iloc[1]))
            except:
                pass
            break

    # === Generic Section Summing ===
    def section_sum_function(start_phrase, label, stop_on_total=True):
        total = count = 0
        for idx, row in df.iterrows():
            if start_phrase.lower() in str(row.iloc[0]).lower():
                j = idx
                while j < len(df):
                    val = str(df.iloc[j, 0]).lower()
                    if any(x in val for x in ["total", "sub total", "grand total"]) and stop_on_total:
                        break
                    cell_val = df.iloc[j, 1]
                    if pd.isna(cell_val) and count > 0:
                        break
                    elif pd.notna(cell_val):
                        try:
                            total += float(cell_val)
                            count += 1
                        except:
                            pass
                    j += 1
                break
        if count > 0:
            insert_values_col4.append(total)
            insert_labels_col3.append(label)

    section_sum_function("reit", "ReIT")
    section_sum_function("invit", "InvIT")
    section_sum_function("foreign securities", "International Equity")

    def total_after_section(section_name, label):
        for idx, row in df.iterrows():
            if str(row.iloc[0]).strip().lower() == section_name.lower():
                for j in range(idx + 1, len(df)):
                    val = str(df.iloc[j, 0]).strip().lower()
                    if val == "total":
                        try:
                            v = float(df.iloc[j, 1])
                            insert_values_col4.append(v)
                            insert_labels_col3.append(label)
                        except:
                            pass
                        return

    total_after_section("Equity & Equity related", "Net Equity")
    total_after_section("Debt Instruments", "Debt")

    def cash_section_sum(term):
        total = count = 0
        for idx, row in df.iterrows():
            if term.lower() in str(row.iloc[0]).lower():
                j = idx
                while j < len(df):
                    val = str(df.iloc[j, 0]).lower()
                    if "total" in val: break
                    cell_val = df.iloc[j, 1]
                    if pd.isna(cell_val) and count > 0: break
                    elif pd.notna(cell_val):
                        try:
                            total += float(cell_val)
                            count += 1
                        except:
                            pass
                    j += 1
                break
        return total

    treps = cash_section_sum("TREPS")
    netrec = cash_section_sum("Net Receivables")

    margin_val = 0
    for idx in range(len(df)):
        if "margin" in str(df.iloc[idx, 0]).lower():
            try:
                margin_val = float(df.iloc[idx, 1])
            except:
                pass
            break

    cash_total = treps + netrec + margin_val
    if cash_total != 0:
        insert_values_col4.append(cash_total)
        insert_labels_col3.append("Cash & Others")

    market_total = 0
    for idx, row in df.iterrows():
        if "market instruments" in str(row.iloc[0]).lower():
            j = idx
            while j < len(df):
                val = str(df.iloc[j, 0]).lower()
                if "total" in val: break
                cell_val = df.iloc[j, 1]
                if pd.notna(cell_val):
                    try:
                        market_total += float(cell_val)
                    except: pass
                j += 1
            break

    for i, lbl in enumerate(insert_labels_col3):
        if lbl == "Debt":
            insert_values_col4[i] += market_total

    for i, lbl in enumerate(insert_labels_col3):
        if lbl == "Net Equity":
            for j, lbl2 in enumerate(insert_labels_col3):
                if lbl2 == "Hedged Equity":
                    insert_values_col4[i] -= abs(insert_values_col4[j])
                    insert_values_col4[j] = abs(insert_values_col4[j])
                    break
            break

    existing_col4 = df.iloc[:, 3].dropna().tolist()
    combined_col4 = insert_values_col4 + existing_col4
    combined_col4 += [None] * (len(df) - len(combined_col4))

    combined_col3 = insert_labels_col3 + [None] * (len(df) - len(insert_labels_col3))

    df.iloc[:, 3] = combined_col4[:len(df)]
    df.iloc[:, 2] = combined_col3[:len(df)]

    return df.iloc[:, [2, 3]]

# ---------------------------- MASTER PIPELINE ----------------------------

def run_master_pipeline(uploaded_files):
    results = {}

    fund_processors = {
        "adityabirla": process_adityabirla,
        # add more like: "axis": process_axis, etc.
    }

    for filename, file_data in uploaded_files.items():
        fund_key = filename.split(".")[0].strip().lower().replace(" ", "")
        processor = fund_processors.get(fund_key)

        if processor:
            try:
                df = processor(file_data.read())
                results[fund_key.title()] = df
            except Exception as e:
                results[fund_key.title()] = f"[ERROR] {fund_key}: {str(e)}"
        else:
            results[fund_key.title()] = f"[SKIPPED] No processor for {fund_key}"

    return results
