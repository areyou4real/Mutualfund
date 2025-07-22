import streamlit as st
import pandas as pd
from io import BytesIO
import re

# ---------------------------- FUND PROCESSORS ----------------------------
def process_adityabirla(file_bytes):
    import pandas as pd
    from io import BytesIO

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

    # === Gold & Silver ===
    for metal, tag in [("gold", "Gold"), ("silver", "Silver")]:
        for idx, row in df.iterrows():
            desc = str(row.iloc[0]).lower()
            if metal in desc:
                try:
                    val = float(row.iloc[1])
                    insert_labels_col3.append(tag)
                    insert_values_col4.append(val)
                except:
                    pass
                break

    # === Section-based tag assignment ===
    def section_total(start_phrase, label, stop_on_total=True):
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
            insert_labels_col3.append(label)
            insert_values_col4.append(total)

    section_total("reit", "ReIT")
    section_total("invit", "InvIT")
    section_total("foreign securities", "International Equity")

    # === Equity & Debt Totals ===
    equity_val = None
    def total_after_section(section_name, label):
        nonlocal equity_val
        for idx, row in df.iterrows():
            if str(row.iloc[0]).strip().lower() == section_name.lower():
                for j in range(idx + 1, len(df)):
                    val = str(df.iloc[j, 0]).strip().lower()
                    if val == "total":
                        try:
                            v = float(df.iloc[j, 1])
                            if label == "Net Equity":
                                equity_val = v
                            insert_labels_col3.append(label)
                            insert_values_col4.append(v)
                        except:
                            pass
                        return

    total_after_section("Equity & Equity related", "Net Equity")
    total_after_section("Debt Instruments", "Debt")

    # === Cash & Others ===
    def cash_sum(term):
        total = count = 0
        for idx, row in df.iterrows():
            if term.lower() in str(row.iloc[0]).lower():
                j = idx
                while j < len(df):
                    val = str(df.iloc[j, 0]).lower()
                    if "total" in val:
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
        return total if count > 0 else 0

    treps = cash_sum("TREPS")
    netrec = cash_sum("Net Receivables")
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
        insert_labels_col3.append("Cash & Others")
        insert_values_col4.append(cash_total)

    # === Market Instruments under Debt ===
    market_total = 0
    for idx, row in df.iterrows():
        if "market instruments" in str(row.iloc[0]).lower():
            j = idx
            while j < len(df):
                val = str(df.iloc[j, 0]).lower()
                if "total" in val:
                    break
                cell_val = df.iloc[j, 1]
                if pd.notna(cell_val):
                    try:
                        market_total += float(cell_val)
                    except:
                        pass
                j += 1
            break
    for i, lbl in enumerate(insert_labels_col3):
        if lbl == "Debt":
            insert_values_col4[i] += market_total

    # === Hedged Equity offset ===
    for i, lbl in enumerate(insert_labels_col3):
        if lbl == "Net Equity":
            for j, lbl2 in enumerate(insert_labels_col3):
                if lbl2 == "Hedged Equity":
                    insert_values_col4[i] -= abs(insert_values_col4[j])
                    insert_values_col4[j] = abs(insert_values_col4[j])
                    break
            break

    # === Final output ===
    return pd.DataFrame({
        "Tag": insert_labels_col3,
        "Final Value": insert_values_col4
    })




def process_axis(file_bytes):
    import pandas as pd
    from io import BytesIO

    df = pd.read_excel(BytesIO(file_bytes), header=None)
    df_filtered = df.iloc[:, [1, 6]].copy()
    df_filtered.columns = ["Category", "Value"]
    df_filtered["Tag"] = None
    df_filtered["Final Value"] = None

    tags = [
        "Hedged Equity", "Net Equity", "Debt", "Gold",
        "Silver", "International Equity", "ReIT/InvIT", "Cash & others"
    ]
    final_values = {tag: 0.0 for tag in tags}

    def get_total_after_match(start_idx, allowed_totals):
        for i in range(start_idx + 1, len(df_filtered)):
            val = str(df_filtered.loc[i, "Category"]).strip().lower()
            if any(val == at for at in allowed_totals):
                try:
                    return float(df_filtered.loc[i, "Value"])
                except:
                    return 0.0
        return 0.0

    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "derivatives" in val.lower():
            final_values["Hedged Equity"] = get_total_after_match(i, ["total"])
            break

    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "equity & equity related" in val.lower():
            final_values["Net Equity"] = get_total_after_match(i, ["total"])
            break

    debt_total = 0.0
    for keyword in ["debt instruments", "money market instruments"]:
        for i, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and keyword in val.lower():
                debt_total += get_total_after_match(i, ["total"])
                break
    final_values["Debt"] = debt_total

    final_values["Net Equity"] -= abs(final_values["Hedged Equity"])
    final_values["Hedged Equity"] = abs(final_values["Hedged Equity"])

    def sum_keyword(keyword):
        total = 0.0
        for i, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and keyword in val.lower():
                try:
                    total += float(df_filtered.loc[i, "Value"])
                except:
                    continue
        return total

    final_values["Gold"] = sum_keyword("gold")
    final_values["Silver"] = sum_keyword("silver")

    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "foreign" in val.lower():
            final_values["International Equity"] = get_total_after_match(i, ["total"])
            break

    reit_invit_total = 0.0
    for keyword in ["reit", "invit"]:
        for i, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and val.lower().startswith(keyword):
                reit_invit_total += get_total_after_match(i, ["total", "sub total"])
                break
    final_values["ReIT/InvIT"] = reit_invit_total

    reverse_val = 0.0
    net_recv_val = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "reverse repo" in val.lower():
            reverse_val = get_total_after_match(i, ["sub total"])
            break
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "net receivables" in val.lower():
            try:
                net_recv_val = float(df_filtered.loc[i, "Value"])
            except:
                net_recv_val = 0.0
            break

    final_values["Cash & others"] = reverse_val + net_recv_val

    summary_df = pd.DataFrame({
        "Category": [None] * len(final_values),
        "Value": [None] * len(final_values),
        "Tag": list(final_values.keys()),
        "Final Value": list(final_values.values())
    })

    final_df = pd.concat([summary_df, df_filtered], ignore_index=True)
    return final_df.iloc[:, [2, 3]]  # Tag, Final Value

def process_baroda(file_bytes):
    import pandas as pd
    from io import BytesIO

    df = pd.read_excel(BytesIO(file_bytes), header=None)
    df_filtered = df.iloc[:, [1, 7]].copy()
    df_filtered.columns = ["Category", "Value"]
    df_filtered["Tag"] = None
    df_filtered["Final Value"] = None

    tags = [
        "Hedged Equity", "Net Equity", "Debt", "Gold",
        "Silver", "International Equity", "ReIT/InvIT", "Cash & others"
    ]
    final_values = {tag: 0.0 for tag in tags}

    def get_total_after_match(start_idx, allowed_totals):
        for i in range(start_idx + 1, len(df_filtered)):
            val = str(df_filtered.loc[i, "Category"]).strip().lower()
            if any(val == at for at in allowed_totals):
                try:
                    return float(df_filtered.loc[i, "Value"])
                except:
                    return 0.0
        return 0.0

    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "derivatives" in val.lower():
            final_values["Hedged Equity"] = get_total_after_match(i, ["total"])
            break

    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "equity & equity related" in val.lower():
            final_values["Net Equity"] = get_total_after_match(i, ["sub total"])
            break

    debt_total = 0.0
    for keyword in ["debt instruments", "money market instruments"]:
        for i, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and keyword in val.lower():
                debt_total += get_total_after_match(i, ["total"])
                break
    final_values["Debt"] = debt_total

    final_values["Net Equity"] -= abs(final_values["Hedged Equity"])
    final_values["Hedged Equity"] = abs(final_values["Hedged Equity"])

    def sum_keyword(keyword):
        total = 0.0
        for i, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and keyword in val.lower():
                try:
                    total += float(df_filtered.loc[i, "Value"])
                except:
                    continue
        return total

    final_values["Gold"] = sum_keyword("gold")
    final_values["Silver"] = sum_keyword("silver")

    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "foreign" in val.lower():
            final_values["International Equity"] = get_total_after_match(i, ["total"])
            break

    reit_invit_total = 0.0
    for keyword in ["reits", "invits"]:
        for i, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and keyword in val.lower():
                reit_invit_total += get_total_after_match(i, ["total", "sub total"])
    final_values["ReIT/InvIT"] = reit_invit_total

    reverse_val = 0.0
    net_recv_val = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "reverse repo" in val.lower():
            reverse_val = get_total_after_match(i, ["sub total"])
            break
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "net receivables" in val.lower():
            try:
                net_recv_val = float(df_filtered.loc[i, "Value"])
            except:
                net_recv_val = 0.0
            break
    final_values["Cash & others"] = reverse_val + net_recv_val

    summary_df = pd.DataFrame({
        "Category": [None] * len(final_values),
        "Value": [None] * len(final_values),
        "Tag": list(final_values.keys()),
        "Final Value": list(final_values.values())
    })

    final_df = pd.concat([summary_df, df_filtered], ignore_index=True)
    return final_df.iloc[:, [2, 3]]  # Tag, Final Value

def process_hdfc(file_bytes):
    import pandas as pd
    from io import BytesIO

    df = pd.read_excel(BytesIO(file_bytes), sheet_name="MY2005", header=None)
    df_filtered = df.iloc[:, [1, 3, 7, 10, 11]]  # columns 2, 4, 8, 11, 12
    df_filtered.columns = ["ISIN/Description", "Name", "Exposure", "Value 1", "Value 2"]
    df_filtered["Tag"] = None
    df_filtered["Value"] = None
    insert_row = 0

    # Portfolio classification section
    portfolio_start = None
    for idx, val in enumerate(df_filtered.iloc[:, 0]):
        if isinstance(val, str) and "portfolio classification" in val.lower():
            portfolio_start = idx
            break

    # Net Equity
    if portfolio_start is not None:
        for j in range(portfolio_start + 1, len(df_filtered)):
            if "equity" in str(df_filtered.iloc[j, 0]).lower():
                df_filtered.at[insert_row, "Value"] = df_filtered.iloc[j, 1]
                df_filtered.at[insert_row, "Tag"] = "Net Equity"
                insert_row += 1
                break

        # Hedged Equity
        for j in range(portfolio_start + 1, len(df_filtered)):
            if "total hedged exposure" in str(df_filtered.iloc[j, 0]).lower():
                df_filtered.at[insert_row, "Value"] = df_filtered.iloc[j, 1]
                df_filtered.at[insert_row, "Tag"] = "Hedged Equity"
                insert_row += 1
                break

    # ReIT + InvIT
    reit_val = invit_val = 0.0
    for j in range(portfolio_start + 1, len(df_filtered)):
        val = str(df_filtered.iloc[j, 0]).lower()
        if "units issued by reit" in val:
            try: reit_val = float(df_filtered.iloc[j, 1])
            except: pass
        if "units issued by invit" in val:
            try: invit_val = float(df_filtered.iloc[j, 1])
            except: pass
    df_filtered.at[insert_row, "Value"] = reit_val + invit_val
    df_filtered.at[insert_row, "Tag"] = "ReIT/InvIT"
    insert_row += 1

    # Cash & Others
    for j in range(portfolio_start + 1, len(df_filtered)):
        if "cash" in str(df_filtered.iloc[j, 0]).lower():
            df_filtered.at[insert_row, "Value"] = df_filtered.iloc[j, 1]
            df_filtered.at[insert_row, "Tag"] = "Cash & Others"
            insert_row += 1
            break

    # Gold
    gold_value = 0.0
    for idx in range(len(df_filtered)):
        name_val = df_filtered.iloc[idx, 1]
        if isinstance(name_val, str) and "gold" in name_val.lower() and "fund" in name_val.lower():
            try: gold_value += float(df_filtered.iloc[idx, 2])
            except: continue
    if gold_value > 0:
        df_filtered.loc[insert_row, "Value"] = gold_value
        df_filtered.loc[insert_row, "Tag"] = "Gold"
        insert_row += 1

    # Silver
    silver_total = 0.0
    for idx in range(len(df_filtered)):
        name_val = df_filtered.iloc[idx, 1]
        if isinstance(name_val, str) and "silver" in name_val.lower():
            for j in range(idx, len(df_filtered)):
                val = df_filtered.iloc[j, 2]
                if pd.isna(val) and j != idx: break
                try: silver_total += float(val)
                except: pass
            break
    df_filtered.at[insert_row, "Value"] = silver_total
    df_filtered.at[insert_row, "Tag"] = "Silver"
    insert_row += 1

    # Debt Instruments
    debt_val = cd_value = 0.0
    for idx, val in enumerate(df_filtered.iloc[:, 0]):
        if isinstance(val, str) and val.strip().lower() == "debt instruments":
            for j in range(idx + 1, len(df_filtered)):
                if str(df_filtered.iloc[j, 0]).strip().lower() == "total":
                    try: debt_val = float(df_filtered.iloc[j, 2])
                    except: pass
                    break
            break
    for j in range(portfolio_start + 1, len(df_filtered)):
        if "cd" in str(df_filtered.iloc[j, 0]).lower():
            try: cd_value = float(df_filtered.iloc[j, 1])
            except: pass
            break
    df_filtered.at[insert_row, "Value"] = debt_val + cd_value
    df_filtered.at[insert_row, "Tag"] = "Debt"
    insert_row += 1

    # International Equity
    intl_value = 0.0
    for idx, val in enumerate(df_filtered.iloc[:, 0]):
        if isinstance(val, str) and val.strip().lower() == "international":
            for j in range(idx + 1, len(df_filtered)):
                if str(df_filtered.iloc[j, 0]).strip().lower() == "total":
                    try: intl_value = float(df_filtered.iloc[j, 2])
                    except: pass
                    break
            break
    df_filtered.at[insert_row, "Value"] = intl_value
    df_filtered.at[insert_row, "Tag"] = "International equity"
    insert_row += 1

    return df_filtered[["Tag", "Value"]].dropna()

def process_hsbc(file_bytes):
    import pandas as pd
    from io import BytesIO

    df = pd.read_excel(BytesIO(file_bytes), header=None)
    df_filtered = df.iloc[:, [0, 5]].copy()
    df_filtered.columns = ["Category", "Value"]
    df_filtered["Tag"] = None
    df_filtered["Final Value"] = None

    def find_total_after(keyword, n=1):
        count = 0
        total = 0.0
        start = None
        for i, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and keyword.lower() in val.lower():
                start = i
                break
        if start is not None:
            for j in range(start + 1, len(df_filtered)):
                val = str(df_filtered.loc[j, "Category"]).strip().lower()
                if val == "total" or val == "sub total":
                    try:
                        val_to_add = float(df_filtered.loc[j, "Value"])
                        if not pd.isna(val_to_add):
                            total += val_to_add
                            count += 1
                    except:
                        pass
                    if count == n:
                        break
        return total

    net_equity_value = find_total_after("Equity & Equity Related Instruments", 1)
    debt_value = find_total_after("Debt Instruments", 3)
    gold_value = find_total_after("Exchange Traded Fund", 1)

    # Cash (TREPS + Net Current Assets)
    treps_value = nca_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str):
            if "treps" in val.lower():
                try:
                    treps_value = float(df_filtered.loc[i, "Value"])
                except: pass
            if "net current assets" in val.lower():
                try:
                    nca_value = float(df_filtered.loc[i, "Value"])
                except: pass
    cash_value = treps_value + nca_value

    # ReIT/InvIT
    reit_invit_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and ("reits" in val.lower() or "invits" in val.lower()):
            reit_invit_value = find_total_after(val, 1)
            break

    # International Equity
    foreign_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "foreign" in val.lower():
            foreign_value = find_total_after(val, 1)
            break

    # Silver
    silver_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "silver" in val.lower():
            try:
                silver_value += float(df_filtered.loc[i, "Value"])
            except:
                pass

    summary_df = pd.DataFrame({
        "Category": [None] * 7,
        "Value": [None] * 7,
        "Tag": [
            "Net Equity", "Debt", "Gold", "Cash",
            "ReIT/InvIT", "International Equity", "Silver"
        ],
        "Final Value": [
            net_equity_value, debt_value, gold_value, cash_value,
            reit_invit_value, foreign_value, silver_value
        ]
    })

    final_df = pd.concat([summary_df, df_filtered], ignore_index=True)
    return final_df[["Tag", "Final Value"]]

def process_icici(file_bytes):
    import pandas as pd
    from io import BytesIO

    df = pd.read_excel(BytesIO(file_bytes), sheet_name='MULTI', header=None)
    df_filtered = df.iloc[:, [1, 7]].copy()
    df_filtered.columns = ["Category", "Value"]
    df_filtered["Tag"] = None
    df_filtered["Final Value"] = None

    final_values = {
        "Debt": 0.0,
        "International Equity": 0.0,
        "ReIT/InvIT": 0.0,
        "Gold": 0.0,
        "Silver": 0.0,
        "Commodity Derivatives": 0.0,
        "Hedged equity": 0.0,
        "Net equity": 0.0,
        "Cash & others": 0.0
    }

    debt_keywords = [
        "Debt Instruments", "Money Market Instruments",
        "Compulsory Convertible Debenture"
    ]
    for keyword in debt_keywords:
        for idx, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and keyword.lower() in val.lower():
                try:
                    final_values["Debt"] += float(df_filtered.loc[idx, "Value"])
                    break
                except:
                    pass

    for idx, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "foreign securities" in val.lower():
            try:
                final_values["International Equity"] += float(df_filtered.loc[idx, "Value"])
            except:
                pass
            break

    reit_invit_keywords = ["reit", "invit"]
    found = set()
    for idx, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str):
            lower_val = val.lower()
            for keyword in reit_invit_keywords:
                if keyword in lower_val and keyword not in found:
                    try:
                        final_values["ReIT/InvIT"] += float(df_filtered.loc[idx, "Value"])
                        found.add(keyword)
                    except:
                        pass
            if len(found) == len(reit_invit_keywords):
                break

    etf_keywords = {"Gold": "gold etf", "Silver": "silver etf"}
    for tag, keyword in etf_keywords.items():
        for idx, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and keyword in val.lower():
                try:
                    final_values[tag] += float(df_filtered.loc[idx, "Value"])
                except:
                    pass
                break

    # Commodity Derivatives
    for idx, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "exchange traded commodity derivatives" in val.lower():
            total = 0.0
            started = False
            for j in range(idx, len(df_filtered)):
                cell_val = df_filtered.loc[j, "Value"]
                if pd.isna(cell_val):
                    if started: break
                    continue
                try:
                    total += float(cell_val)
                    started = True
                except:
                    if started: break
            final_values["Commodity Derivatives"] += total
            break

    # Hedged Equity
    for idx, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "stock / index futures" in val.lower():
            total = 0.0
            started = False
            for j in range(idx, len(df_filtered)):
                cell_val = df_filtered.loc[j, "Value"]
                if pd.isna(cell_val):
                    if started: break
                    continue
                try:
                    total += float(cell_val)
                    started = True
                except:
                    if started: break
            final_values["Hedged equity"] += total
            break

    # Net Equity
    equity_start = None
    listed_val = None
    for idx, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "equity" in val.lower():
            equity_start = idx
            break
    if equity_start is not None:
        for j in range(equity_start + 1, len(df_filtered)):
            val = df_filtered.loc[j, "Category"]
            if isinstance(val, str) and "listed" in val.lower():
                try:
                    listed_val = float(df_filtered.loc[j, "Value"])
                except:
                    pass
                break
    if listed_val is not None:
        final_values["Net equity"] = listed_val - abs(final_values["Hedged equity"])
        final_values["Hedged equity"] = abs(final_values["Hedged equity"])

    # Cash & Others
    cash_keywords = ["TREPS", "Net current assets"]
    cash_total = 0.0
    found_cash = set()
    for idx, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str):
            lower_val = val.lower()
            for keyword in cash_keywords:
                if keyword.lower() in lower_val and keyword not in found_cash:
                    try:
                        cash_total += float(df_filtered.loc[idx, "Value"])
                        found_cash.add(keyword)
                    except:
                        pass
        if len(found_cash) == len(cash_keywords):
            break
    final_values["Cash & others"] = cash_total

    summary_rows = pd.DataFrame({
        "Category": [None] * len(final_values),
        "Value": [None] * len(final_values),
        "Tag": list(final_values.keys()),
        "Final Value": list(final_values.values())
    })

    df_combined = pd.concat([summary_rows, df_filtered], ignore_index=True)
    return df_combined[["Tag", "Final Value"]]

def process_mahindra(file_bytes):
    import pandas as pd
    from io import BytesIO

    df = pd.read_excel(BytesIO(file_bytes), sheet_name="MMF23", header=None)
    df_filtered = df.iloc[:, [1, 6]].copy()
    df_filtered.columns = ["Category", "Value"]
    df_filtered["Tag"] = None
    df_filtered["Final Value"] = None

    def find_exact_total_after(keyword, first_word_only=False, allow_sub_total=False):
        start = None
        for i, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str):
                val_lower = val.lower().strip()
                if first_word_only:
                    if val_lower.startswith(keyword.lower()):
                        start = i
                        break
                elif keyword.lower() in val_lower:
                    start = i
                    break
        if start is not None:
            for j in range(start + 1, len(df_filtered)):
                val = str(df_filtered.loc[j, "Category"]).strip().lower()
                if val == "total" or (allow_sub_total and "total" in val):
                    try:
                        return float(df_filtered.loc[j, "Value"])
                    except:
                        return 0.0
        return 0.0

    # Major components
    net_equity_value = find_exact_total_after("Equity & Equity related")
    debt_value = find_exact_total_after("Debt instruments")
    reits_value = find_exact_total_after("reits")
    invits_value = find_exact_total_after("invits")
    treps_value = find_exact_total_after("treps")

    # Gold and Silver
    gold_value = silver_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str):
            if "gold" in val.lower():
                try:
                    v = float(df_filtered.loc[i, "Value"])
                    if not pd.isna(v): gold_value += v
                except: pass
            if "silver" in val.lower():
                try:
                    v = float(df_filtered.loc[i, "Value"])
                    if not pd.isna(v): silver_value += v
                except: pass

    # Net Receivables
    net_recv_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "net receivables" in val.lower():
            try:
                net_recv_value = float(df_filtered.loc[i, "Value"])
            except:
                pass
            break
    cash_value = treps_value + net_recv_value

    # Foreign = International Equity
    foreign_value = find_exact_total_after("foreign", first_word_only=True)

    # Derivatives = Hedged Equity
    deriv_value = find_exact_total_after("derivatives", first_word_only=True)

    summary_df = pd.DataFrame({
        "Category": [None] * 9,
        "Value": [None] * 9,
        "Tag": [
            "Equity", "Debt", "ReITs", "InvITs",
            "Gold", "Silver", "Cash",
            "International Equity", "Hedged Equity"
        ],
        "Final Value": [
            net_equity_value, debt_value, reits_value, invits_value,
            gold_value, silver_value, cash_value,
            foreign_value, deriv_value
        ]
    })

    final_df = pd.concat([summary_df, df_filtered], ignore_index=True)
    return final_df[["Tag", "Final Value"]]

def process_mirae(file_bytes):
    import pandas as pd
    from io import BytesIO

    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    df_filtered = df.iloc[:, [1, 6]].copy()
    df_filtered.columns = ["Category", "Value"]
    df_filtered["Tag"] = None
    df_filtered["Final Value"] = None

    def is_valid_number(val):
        if isinstance(val, (int, float)): return True
        if isinstance(val, str):
            val = val.strip().lower()
            if val in ["", "nil", "na", "n.a.", "-", "--"]:
                return False
            try: float(val); return True
            except: return False
        return False

    def find_exact_total_after(keyword):
        start = None
        for i, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and keyword.lower() in val.lower():
                start = i
                break
        if start is not None:
            for j in range(start + 1, len(df_filtered)):
                cat = str(df_filtered.loc[j, "Category"]).strip().lower()
                value = df_filtered.loc[j, "Value"]
                if cat == "total" and is_valid_number(value):
                    return float(value)
        return 0.0

    def find_sub_total_after(keyword):
        start = None
        for i, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and keyword.lower() in val.lower():
                start = i
                break
        if start is not None:
            for j in range(start + 1, len(df_filtered)):
                cat = str(df_filtered.loc[j, "Category"]).strip().lower()
                value = df_filtered.loc[j, "Value"]
                if cat == "sub total" and is_valid_number(value):
                    return float(value)
        return 0.0

    # Extraction logic
    net_equity_value = find_exact_total_after("Equity & Equity related")
    debt_instr = find_exact_total_after("Debt instruments")
    real_estate_sub = find_sub_total_after("Real Estate Investment Trust")
    debt_value = 0.0
    if is_valid_number(debt_instr): debt_value += debt_instr
    if is_valid_number(real_estate_sub): debt_value += real_estate_sub

    reits_value = find_exact_total_after("reits")
    invits_value = find_exact_total_after("invits")
    treps_value = find_sub_total_after("treps")

    gold_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "gold" in val.lower():
            raw_val = df_filtered.loc[i, "Value"]
            if pd.notna(raw_val) and is_valid_number(raw_val):
                gold_value += float(raw_val)

    silver_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "silver" in val.lower():
            raw_val = df_filtered.loc[i, "Value"]
            if pd.notna(raw_val) and is_valid_number(raw_val):
                silver_value += float(raw_val)

    net_recv_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "net receivables" in val.lower():
            if is_valid_number(df_filtered.loc[i, "Value"]):
                net_recv_value = float(df_filtered.loc[i, "Value"])
            break

    cash_value = treps_value + net_recv_value

    foreign_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and val.lower().startswith("foreign"):
            for j in range(i + 1, len(df_filtered)):
                if str(df_filtered.loc[j, "Category"]).strip().lower() == "total" and is_valid_number(df_filtered.loc[j, "Value"]):
                    foreign_value = float(df_filtered.loc[j, "Value"])
                    break
            break

    deriv_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and val.lower().startswith("derivatives"):
            for j in range(i + 1, len(df_filtered)):
                if str(df_filtered.loc[j, "Category"]).strip().lower() == "total" and is_valid_number(df_filtered.loc[j, "Value"]):
                    deriv_value = float(df_filtered.loc[j, "Value"])
                    break
            break

    net_equity_value -= abs(deriv_value)
    deriv_value = abs(deriv_value)

    summary_df = pd.DataFrame({
        "Category": [None] * 9,
        "Value": [None] * 9,
        "Tag": [
            "Equity", "Debt", "ReITs", "InvITs",
            "Gold", "Silver", "Cash",
            "International Equity", "Hedged Equity"
        ],
        "Final Value": [
            net_equity_value, debt_value, reits_value, invits_value,
            gold_value, silver_value, cash_value,
            foreign_value, deriv_value
        ]
    })

    final_df = pd.concat([summary_df, df_filtered], ignore_index=True)
    return final_df[["Tag", "Final Value"]]

def process_shriram(file_bytes):
    import pandas as pd
    from io import BytesIO

    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    df_filtered = df.iloc[:, [1, 6]].copy()
    df_filtered.columns = ["Category", "Value"]
    df_filtered["Tag"] = None
    df_filtered["Final Value"] = None

    def is_valid_number(val):
        if isinstance(val, (int, float)): return True
        if isinstance(val, str):
            val = val.strip().lower()
            if val in ["", "nil", "na", "n.a.", "-", "--"]:
                return False
            try: float(val); return True
            except: return False
        return False

    def find_exact_total_after(keyword):
        for i, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and keyword.lower() in val.lower():
                for j in range(i + 1, len(df_filtered)):
                    cat = str(df_filtered.loc[j, "Category"]).strip().lower()
                    value = df_filtered.loc[j, "Value"]
                    if cat == "total" and is_valid_number(value):
                        return float(value)
        return 0.0

    def find_sub_total_after(keyword):
        for i, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and keyword.lower() in val.lower():
                for j in range(i + 1, len(df_filtered)):
                    cat = str(df_filtered.loc[j, "Category"]).strip().lower()
                    value = df_filtered.loc[j, "Value"]
                    if cat == "sub total" and is_valid_number(value):
                        return float(value)
        return 0.0

    net_equity_value = find_exact_total_after("Equity & Equity related")
    debt_instr = find_exact_total_after("Debt instruments")
    real_estate_sub = find_sub_total_after("Real Estate Investment Trust")
    debt_value = debt_instr + real_estate_sub

    reits_value = find_exact_total_after("reits")
    invits_value = find_exact_total_after("invits")
    treps_value = find_sub_total_after("treps")

    gold_value = silver_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str):
            raw_val = df_filtered.loc[i, "Value"]
            if pd.notna(raw_val) and is_valid_number(raw_val):
                if "gold" in val.lower():
                    gold_value += float(raw_val)
                elif "silver" in val.lower():
                    silver_value += float(raw_val)

    net_recv_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "net receivables" in val.lower():
            if is_valid_number(df_filtered.loc[i, "Value"]):
                net_recv_value = float(df_filtered.loc[i, "Value"])
            break

    cash_value = treps_value + net_recv_value

    foreign_value = deriv_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str):
            if val.lower().startswith("foreign"):
                for j in range(i + 1, len(df_filtered)):
                    if str(df_filtered.loc[j, "Category"]).strip().lower() == "total" and is_valid_number(df_filtered.loc[j, "Value"]):
                        foreign_value = float(df_filtered.loc[j, "Value"])
                        break
            if val.lower().startswith("derivatives"):
                for j in range(i + 1, len(df_filtered)):
                    if str(df_filtered.loc[j, "Category"]).strip().lower() == "total" and is_valid_number(df_filtered.loc[j, "Value"]):
                        deriv_value = float(df_filtered.loc[j, "Value"])
                        break

    net_equity_value -= abs(deriv_value)
    deriv_value = abs(deriv_value)

    summary_df = pd.DataFrame({
        "Category": [None] * 9,
        "Value": [None] * 9,
        "Tag": [
            "Equity", "Debt", "ReITs", "InvITs",
            "Gold", "Silver", "Cash",
            "International Equity", "Hedged Equity"
        ],
        "Final Value": [
            net_equity_value, debt_value, reits_value, invits_value,
            gold_value, silver_value, cash_value,
            foreign_value, deriv_value
        ]
    })

    final_df = pd.concat([summary_df, df_filtered], ignore_index=True)
    return final_df[["Tag", "Final Value"]]

def process_sundaram(file_bytes):
    import pandas as pd
    from io import BytesIO

    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    df_filtered = df.iloc[:, [2, 6]].copy()
    df_filtered.columns = ["Category", "Value"]
    df_filtered["Tag"] = None
    df_filtered["Final Value"] = None

    def is_valid_number(val):
        if isinstance(val, (int, float)): return True
        if isinstance(val, str):
            val = val.strip().lower()
            if val in ["", "nil", "na", "n.a.", "-", "--"]: return False
            try: float(val); return True
            except: return False
        return False

    def find_next_summary_after(keyword, target="total"):
        for i, val in enumerate(df_filtered["Category"]):
            if isinstance(val, str) and keyword.lower() in val.lower():
                for j in range(i + 1, len(df_filtered)):
                    cat = str(df_filtered.loc[j, "Category"]).strip().lower()
                    value = df_filtered.loc[j, "Value"]
                    if cat == target and is_valid_number(value):
                        return float(value)
        return 0.0

    equity_value = find_next_summary_after("Equity & Equity related", target="sub total")

    # Debt = Total for Debt Instruments + Treasury Bills
    debt_main = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and "Total for Debt Instruments" in val:
            if is_valid_number(df_filtered.loc[i, "Value"]):
                debt_main = float(df_filtered.loc[i, "Value"])
            break
    treasury_value = find_next_summary_after("Treasury Bills", target="sub total")
    debt_value = debt_main + treasury_value

    reits_value = find_next_summary_after("reits")
    invits_value = find_next_summary_after("invits")
    treps_value = find_next_summary_after("Treps", target="sub total")

    # Gold and Silver
    gold_value = silver_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str):
            v = df_filtered.loc[i, "Value"]
            if "gold" in val.lower() and is_valid_number(v): gold_value += float(v)
            elif "silver" in val.lower() and is_valid_number(v): silver_value += float(v)

    # Margin Money and Cash and Other
    margin_val = cashother_val = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str):
            if "margin money" in val.lower() and is_valid_number(df_filtered.loc[i, "Value"]):
                margin_val = abs(float(df_filtered.loc[i, "Value"]))
            if "cash and other" in val.lower() and is_valid_number(df_filtered.loc[i, "Value"]):
                cashother_val = abs(float(df_filtered.loc[i, "Value"]))
    cash_value = abs(treps_value) - (margin_val + cashother_val)

    # Derivatives and Foreign
    deriv_value = find_next_summary_after("derivative", target="sub total")
    foreign_value = 0.0
    for i, val in enumerate(df_filtered["Category"]):
        if isinstance(val, str) and val.lower().startswith("foreign"):
            for j in range(i + 1, len(df_filtered)):
                if str(df_filtered.loc[j, "Category"]).strip().lower() == "total" and is_valid_number(df_filtered.loc[j, "Value"]):
                    foreign_value = float(df_filtered.loc[j, "Value"])
                    break
            break

    # Adjust equity
    adjusted_equity = equity_value - abs(deriv_value)
    deriv_value = abs(deriv_value)

    summary_df = pd.DataFrame({
        "Category": [None] * 9,
        "Value": [None] * 9,
        "Tag": [
            "Equity", "Debt", "ReITs", "InvITs",
            "Gold", "Silver", "Cash",
            "International Equity", "Hedged Equity"
        ],
        "Final Value": [
            adjusted_equity, debt_value, reits_value, invits_value,
            gold_value, silver_value, cash_value,
            foreign_value, deriv_value
        ]
    })

    final_df = pd.concat([summary_df, df_filtered], ignore_index=True)
    return final_df[["Tag", "Final Value"]]

# ---------------------------- NAME NORMALIZATION ----------------------------
def normalize_name(filename):
    name = re.sub(r"\s\(\d+\)", "", filename)  # removes (1), (2), etc.
    return name.lower().split('.')[0].replace(" ", "_")

# ---------------------------- PROCESSOR MAPPING ----------------------------

fund_processors = {
    "adityabirla": process_adityabirla,
     "axis": process_axis,
     "baroda": process_baroda,
     "hdfc": process_hdfc,
     "hsbc": process_hsbc,
     "icici": process_icici,
     "mahindra": process_mahindra,
     "mirae": process_mirae,
     "shriram": process_shriram,
    "sundaram": process_sundaram
}

def match_processor_key(name_key):
    for key in fund_processors:
        if name_key == key or name_key.replace("_", "") == key.replace("_", ""):
            return fund_processors[key]
    return None

# ---------------------------- MASTER PIPELINE ----------------------------
def run_master_pipeline(uploaded_files):
    output_dfs = {}
    for file_name, file_bytes in uploaded_files.items():
        name_key = normalize_name(file_name)
        processor = match_processor_key(name_key)

        if processor:
            try:
                df = processor(file_bytes)
                if isinstance(df, pd.DataFrame):
                    output_dfs[name_key] = df
                else:
                    output_dfs[name_key] = "Returned object is not a DataFrame"
            except Exception as e:
                output_dfs[name_key] = f"Error: {str(e)}"
        else:
            output_dfs[name_key] = "No matching processor found"
    return output_dfs

# ---------------------------- STREAMLIT UI ----------------------------
from streamlit.components.v1 import html

st.set_page_config(page_title="Mutual Fund Allocation Generator", layout="centered")

st.markdown(
    """
    <style>
.stApp {
    background-color: #121212;
    color: #e0e0e0;
    font-family: 'Segoe UI', sans-serif;
    padding: 2rem 1rem;
    max-width: 960px;
    margin: auto;
}

.title {
    font-size: 2.8rem;
    font-weight: 600;
    text-align: center;
    margin-top: 1rem;
    color: #ffffff;
    margin-bottom: 0.25rem;
}

.subtitle {
    font-size: 1.2rem;
    text-align: center;
    color: #aaaaaa;
    margin-bottom: 2rem;
}

.stButton>button {
    background-color: #1e88e5;
    color: #ffffff;
    border-radius: 6px;
    border: none;
    padding: 0.6rem 1.2rem;
    font-size: 0.95rem;
    transition: background-color 0.2s ease;
}

.stButton>button:hover {
    background-color: #1565c0;
}

.css-1kyxreq.edgvbvh3 {
    background-color: #1e1e1e;
    border: 1px dashed #444;
    border-radius: 8px;
    padding: 1.25rem;
    color: #ccc;
    margin-bottom: 2rem;
}

.block-container {
    padding-top: 2rem;
}
    
@keyframes fadeSlideUp {
    0% {opacity: 0; transform: translateY(20px);}
    100% {opacity: 1; transform: translateY(0);}
}

.fade-in {
    animation: fadeSlideUp 0.6s ease-out forwards;
    opacity: 0;
}
</style>
    """,
    unsafe_allow_html=True
)

st.markdown('<div class="title fade-in">🧾 Mutual Fund Allocation Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle fade-in">Upload one or more mutual fund Excel files. The app will detect the allocations and return a cleaned summary.</div>', unsafe_allow_html=True)

st.markdown('<div class="fade-in">', unsafe_allow_html=True)
uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    file_dict = {file.name: file.read() for file in uploaded_files}
    results = run_master_pipeline(file_dict)

    valid_results = {k: v for k, v in results.items() if isinstance(v, pd.DataFrame)}
    error_results = {k: v for k, v in results.items() if not isinstance(v, pd.DataFrame)}

    if error_results:
        st.subheader("❌ Errors Detected")
        for name, error in error_results.items():
            st.error(f"{name}: {error}")

    if valid_results:
        st.subheader("✅ Fund Allocation Summaries")
        for name, df in valid_results.items():
            with st.expander(f"📁 {name.title()} Allocation Summary"):
                st.dataframe(df.style.format({"Final Value": "{:.2f}"}))

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for name, df in valid_results.items():
                df.to_excel(writer, sheet_name=name[:31], index=False)
        output.seek(0)
    st.markdown('</div>', unsafe_allow_html=True)
    st.download_button("📥 Download All Results", output, file_name="Allocation_Output.xlsx", use_container_width=True)
    else:
        st.info("🔎 No valid dataframes to display or download.")
