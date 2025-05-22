import streamlit as st
import io
import re
import tempfile
import os
import pandas as pd
import pdfplumber
import fitz  # PyMuPDF
import camelot
from pypdf import PdfReader
from io import BytesIO

##############################################################################
# CUSTOM CSS (Teal header row with white text, black body text, centered header)
##############################################################################
custom_css = """
<style>
/* Additional global styling if needed */

/* Make the Pandas-generated .dataframe table headers teal with white text, centered text */
.dataframe thead tr th {
    background-color: #004a5f !important;
    color: #ffffff !important;
    text-align: center !important;
}

/* Force the table body text to be black (so it doesn't vanish) */
.dataframe tbody tr td {
    color: #000000 !important;
}
</style>
"""

##############################################################################
# HELPER FUNCTION TO MAKE TABLE CELLS EDITABLE
##############################################################################
def make_table_cells_editable(html_str: str) -> str:
    """
    Insert contenteditable="true" into each <td> tag so the user can type/edit values.
    """
    pattern = r'(<td)([^>]*>)'
    replace = r'<td contenteditable="true"\2'
    return re.sub(pattern, replace, html_str, flags=re.IGNORECASE)

##############################################################################
# HELPER: CLEAN UNWANTED TEXT FROM POLICY FORM DESCRIPTION
##############################################################################
def clean_description(description: str) -> str:
    """
    Removes extra text from the 'Description' if it contains certain stop phrases.
    We cut off everything from that phrase onward.

    Add or remove stop phrases as needed to handle any unwanted text.
    """
    stop_phrases = [
        "your payment includes",
        "page 1 of 1",
        "commercial automobile ca pn 83 36 tx 07 19 notice",
        "texas motor vehicle crime prevention authority fee",
        "notice -",
        "authority fee · auto burglary,",
        "auto burglary, theft and fraud prevention;",
        "by law, we send this fee to the motor vehicle crime prevention authority",
        "criminal justice efforts;",
        "ca pn 83 36 tx 07 19",
        "commercial automobile · ·",
        "trauma care and emergency medical services for victims of accidents due to traffic offenses."
    ]
    desc_lower = description.lower()
    for phrase in stop_phrases:
        idx = desc_lower.find(phrase)
        if idx != -1:
            # Cut off everything from that phrase onward
            return description[:idx].strip()
    return description

##############################################################################
# EXTRACTION FUNCTIONS
##############################################################################

def format_currency(value_str):
    try:
        cleaned = value_str.replace("$", "").replace(",", "").strip()
        val = float(cleaned)
        return f"${val:,.2f}"
    except:
        return value_str

def extract_table1_pypdf(pdf_data):
    reader = PdfReader(io.BytesIO(pdf_data))
    full_text = ""
    for page in reader.pages:
        text = page.extract_text() or ""
        full_text += text + "\n"
    lines = full_text.splitlines()
    start_index = None
    for i, line in enumerate(lines):
        if "commercial auto coverages premium" in line.lower():
            start_index = i
            break
    if start_index is None:
        return []
    coverage_rows = []
    total_quote_row = None
    for line in lines[start_index + 1:]:
        if "schedule of coverages and covered autos" in line.lower():
            break
        if "$" in line:
            parts = line.split("$")
            if len(parts) >= 2:
                coverage_name = parts[0].strip()
                raw_amount = parts[1].strip()
                premium_value = format_currency(raw_amount)
                if "total quote premium" in coverage_name.lower():
                    total_quote_row = {"Coverage": "Total Quote Premium", "Premium": premium_value}
                else:
                    coverage_rows.append({"Coverage": coverage_name, "Premium": premium_value})
    if total_quote_row:
        coverage_rows.append(total_quote_row)
    return coverage_rows

def extract_text_pymupdf(pdf_data):
    doc = fitz.open(stream=pdf_data, filetype="pdf")
    text_lines = []
    for page in doc:
        text_lines.extend(page.get_text().splitlines())
    doc.close()
    return text_lines


def extract_table2_pymupdf(pdf_data):
    """
    Robust parser for the **Schedule of Coverages and Covered Autos** table.

    The main enhancement is a smarter discrimination between *Covered Autos*
    lists like "7, 8" and monetary figures like "8,818" that appear in the
    *Premium* column but, in raw text, are formatted without the leading "$".
    The heuristic we use:

    ▸ A list of covered‑auto symbols always consists of 1‑ or 2‑digit tokens
      separated by commas, **each token length ≤ 2** (e.g. "7,8", "1, 2, 9").
    ▸ If a comma‑separated string has **any token ≥ 3 digits** (e.g. "8,818"),
      we treat it as a monetary value (Premium) rather than a symbol list.

    In addition, tails of the form “See Schedule 7, 8” are again split so that
    “See Schedule” moves to the Limits column and the symbols “7,8” (now
    correctly recognised) move to Covered Autos.
    """

    import re

    # ------------------------------------------------------------------
    # Utilities
    # ------------------------------------------------------------------
    covered_autos_pat = re.compile(r'^[\d\s,]+$')
    currency_pat      = re.compile(r'^\d{1,3}(?:,\d{3})*(?:\.\d+)?$')

    def looks_like_covered_autos(txt: str) -> bool:
        """
        Heuristic: Every token (split by comma) is 1‑2 digits. Otherwise false.
        """
        tokens = [t.strip() for t in txt.split(',') if t.strip()]
        if not tokens:
            return False
        return all(tok.isdigit() and 1 <= len(tok) <= 2 for tok in tokens)

    # ------------------------------------------------------------------
    # 1) Extract raw text lines from PyMuPDF
    # ------------------------------------------------------------------
    lines = extract_text_pymupdf(pdf_data)

    # Locate the section heading “Schedule of Coverages and Covered Autos”
    try:
        start_idx = next(i for i,l in enumerate(lines)
                         if 'schedule of coverages and covered autos' in l.lower())
    except StopIteration:
        return []

    # Collect until we hit the next major heading or an empty separator line
    table_lines = []
    for ln in lines[start_idx + 1:]:
        low = ln.lower()
        if 'schedule of covered autos you own' in low or 'endorsements' in low:
            break
        table_lines.append(ln.strip())

    # Remove header row if present
    while table_lines and table_lines[0].lower().startswith('coverages'):
        table_lines.pop(0)

    # ------------------------------------------------------------------
    # 2) Iterate through the physical lines, assembling logical rows
    # ------------------------------------------------------------------
    rows = []
    i, n = 0, len(table_lines)
    while i < n:
        # --------------------------------------------------------------
        # A) Coverage / Description (capable of spanning multiple lines)
        # --------------------------------------------------------------
        cov_parts = []
        while i < n and not table_lines[i].startswith('$') \
                   and not table_lines[i].isdigit() \
                   and not covered_autos_pat.match(table_lines[i]):
            cov_parts.append(table_lines[i])
            i += 1
        coverage_str = ' '.join(cov_parts).strip()

        # --------------------------------------------------------------
        # B) Limits
        # --------------------------------------------------------------
        limits_str = ''
        if i < n and (table_lines[i].startswith('$') or
                      table_lines[i].lower().startswith('see schedule') or
                      table_lines[i] == '$'):
            limits_str = table_lines[i]
            i += 1

        # --------------------------------------------------------------
        # C) Covered Autos or Premium (needs heuristic)
        # --------------------------------------------------------------
        covered_autos_str = ''
        premium_str = ''
        if i < n and covered_autos_pat.match(table_lines[i]):
            token = table_lines[i].replace(' ', '')
            if looks_like_covered_autos(token):
                covered_autos_str = token
            else:
                # treat as premium – attach leading $
                premium_str = f'${token}'
            i += 1

        # --------------------------------------------------------------
        # D) Premium (if not determined yet)
        # --------------------------------------------------------------
        if not premium_str and i < n:
            peek = table_lines[i]
            # Variant 1: lone "$" followed by digits
            if peek == '$' and i+1 < n and currency_pat.match(table_lines[i+1]):
                premium_str = '$' + table_lines[i+1]
                i += 2
            # Variant 2: starts with "$"
            elif peek.startswith('$'):
                premium_str = peek
                i += 1
            # Variant 3: bare digits but clearly monetary (≥ 4 digits or ≥ 1000)
            elif currency_pat.match(peek) and len(peek.replace(',', '')) >= 4:
                premium_str = '$' + peek
                i += 1

        # --------------------------------------------------------------
        # Post‑processing for “See Schedule 7, 8” patterns in coverage_str
        # --------------------------------------------------------------
        if 'see schedule' in coverage_str.lower():
            m = re.search(r'\bsee schedule\b\s*([\d\s,]+)$', coverage_str, flags=re.I)
            if m:
                tail = m.group(1).replace(' ', '')
                coverage_str = coverage_str[:m.start()].rstrip(' ,;')
                if looks_like_covered_autos(tail):
                    covered_autos_str = tail
                else:
                    premium_str = '$' + tail  # unlikely but safe‑guard
                if not limits_str:
                    limits_str = 'See Schedule'

        # Normalise strings
        limits_str = limits_str.strip()
        covered_autos_str = covered_autos_str.strip()
        premium_str = premium_str.strip()

        rows.append({
            'Coverages, Limits & Deductibles': coverage_str,
            'Limits'                         : limits_str,
            'Covered Autos'                  : covered_autos_str,
            'Premium'                        : premium_str,
        })

    return rows
def extract_premium_details_pypdf(pdf_data):
    # (Original PyPDF-based approach; not used in final display)
    import re
    from pypdf import PdfReader
    def parse_coverage_value(val):
        val = val.strip()
        if not val or val == "$":
            return "N"
        m = re.search(r'(\d[\d,\.]*)', val)
        if m:
            digits_str = m.group(1).replace(",", "")
            return f"${digits_str}"
        return "N"
    reader = PdfReader(io.BytesIO(pdf_data))
    full_text = ""
    for page in reader.pages:
        txt = page.extract_text() or ""
        full_text += txt + "\n"
    lines = full_text.splitlines()
    premiums_indices = [i for i, line in enumerate(lines) if "PREMIUMS" in line.upper()]
    results = {}
    def merge_lines(raw_lines):
        merged = []
        skip_next = False
        for idx in range(len(raw_lines)):
            if skip_next:
                skip_next = False
                continue
            line = raw_lines[idx]
            if line.count('$') < 3 and (idx + 1) < len(raw_lines):
                merged.append(line + raw_lines[idx+1])
                skip_next = True
            else:
                merged.append(line)
        return merged
    end_markers = {
        "SCHEDULE OF LOSS PAYEES",
        "SCHEDULE OF HIRED OR BORROWED COVERED AUTO COVERAGE AND PREMIUMS"
    }
    def parse_line(line_str):
        match = re.match(r'^\s*(\d+)\s*\$', line_str)
        if not match:
            return None
        veh_no_str = match.group(1)
        remainder = line_str[len(match.group(0)):]
        raw_tokens = remainder.split('$')
        tokens = [t.strip() for t in raw_tokens]
        pip_val     = parse_coverage_value(tokens[1] if len(tokens) > 1 else "")
        med_pay_val = parse_coverage_value(tokens[6] if len(tokens) > 6 else "")
        um_val      = parse_coverage_value(tokens[8] if len(tokens) > 8 else "")
        uim_val     = parse_coverage_value(tokens[9] if len(tokens) > 9 else "")
        return veh_no_str, {"PIP": pip_val, "Med Pay": med_pay_val, "UM": um_val, "UIM": uim_val}
    for start_idx in premiums_indices:
        block = []
        j = start_idx + 1
        while j < len(lines):
            up_line = lines[j].strip().upper()
            if up_line in end_markers:
                break
            block.append(lines[j])
            j += 1
        merged_block = merge_lines(block)
        for bline in merged_block:
            parsed = parse_line(bline.strip())
            if parsed:
                veh_no, cdict = parsed
                results[veh_no] = cdict
    return results

def extract_premium_pdfplumber_for_table4(pdf_data):
    import pdfplumber
    from io import BytesIO
    import re
    text = ""
    with pdfplumber.open(BytesIO(pdf_data)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text += page_text + "\n"
    lines = text.splitlines()
    start_idx = None
    for i, line in enumerate(lines):
        if "PREMIUMS" in line.upper():
            start_idx = i
            break
    if start_idx is None:
        return {}
    premium_lines = []
    for line in lines[start_idx:]:
        if "SCHEDULE OF LOSS PAYEES" in line.upper():
            break
        premium_lines.append(line)
    if len(premium_lines) < 3:
        return {}
    data_lines = premium_lines[2:]
    result = {}
    for dline in data_lines:
        if not dline.strip():
            continue
        parts = dline.split()
        merged = []
        i = 0
        while i < len(parts):
            if parts[i] == "$" and i+1 < len(parts) and re.match(r'^\d[\d,\,\.]*$', parts[i+1]):
                merged.append("$" + parts[i+1])
                i += 2
            else:
                merged.append(parts[i])
                i += 1
        if len(merged) < 12:
            continue
        if not merged[0].isdigit():
            continue
        veh_no = merged[0]
        pip_val = merged[2] if len(merged) > 2 else ""
        med_pay_val = merged[7] if len(merged) > 7 else ""
        um_val = merged[9] if len(merged) > 9 else ""
        uim_val = merged[10] if len(merged) > 10 else ""
        result[veh_no] = {"PIP": pip_val, "Med Pay": med_pay_val, "UM": um_val, "UIM": uim_val}
    return result

def extract_deductibles_pypdf(pdf_data):
    reader = PdfReader(io.BytesIO(pdf_data))
    full_text = ""
    for page in reader.pages:
        full_text += page.extract_text() or "" + "\n"
    lines = full_text.splitlines()
    start_idx = None
    for i, line in enumerate(lines):
        if "Premium Deductibles".lower() in line.lower():
            start_idx = i
            break
    if start_idx is None:
        return {}
    header_found = False
    for i in range(start_idx, len(lines)):
        if "Loss Coll".lower() in lines[i].lower():
            header_found = True
            start_data_idx = i + 1
            break
    if not header_found:
        return {}
    deductibles = {}
    for i in range(start_data_idx, len(lines)):
        if "SCHEDULE OF LOSS PAYEES".lower() in lines[i].lower():
            break
        stripped = lines[i].strip()
        if not stripped:
            continue
        if not re.match(r'^\d+', stripped):
            continue
        tokens = stripped.split()
        veh_no = tokens[0]
        comp_deduct = tokens[1] if len(tokens) > 1 else ""
        coll_deduct = tokens[2] if len(tokens) > 2 else ""
        deductibles[veh_no] = {"Comp Deductible": comp_deduct, "Collision Deductible": coll_deduct}
    return deductibles

def looks_like_vin(value: str) -> bool:
    val = value.strip()
    if len(val) < 8:
        return False
    return any(ch.isdigit() for ch in val) and any(ch.isalpha() for ch in val)

def find_vin_in_text(text: str) -> str:
    upper_text = text.upper()
    cleaned = re.sub(r"[^A-HJ-NPR-Z0-9]", "", upper_text)
    if 15 <= len(cleaned) <= 17:
        return cleaned
    return ""

def extract_state_territory_from_pymupdf(pdf_data):
    doc = fitz.open(stream=pdf_data, filetype="pdf")
    text_lines = []
    for page in doc:
        text_lines.extend(page.get_text().splitlines())
    doc.close()
    tokens = []
    for line in text_lines:
        tokens.extend(line.split())
    results = []
    for i, token in enumerate(tokens):
        if token.lower() == "terr":
            state = None
            territory = None
            if i > 0:
                possible_state = tokens[i-1]
                if re.match(r'^[A-Z]{2}$', possible_state):
                    state = possible_state
            if i+1 < len(tokens):
                possible_territory = tokens[i+1]
                if possible_territory.isdigit():
                    territory = possible_territory.zfill(3)
            if state and territory:
                results.append((state, territory))
    return results

def merge_classification_and_territory(veh_df, tables, pdf_data):
    import re
    from pdfminer.high_level import extract_text as pdfminer_extract_text
    classification_rows = []

    for t in tables:
        df = t.df.copy()
        df = df.reset_index(drop=True)
        df.columns = pd.Series(df.columns.astype(str))
        df = df.loc[:, ~df.columns.duplicated()]

        if df.shape[0] < 2:
            continue

        row0_txt = " ".join(str(x).lower() for x in df.iloc[0])
        if ("classification" in row0_txt) or ("territory (principal garage location)" in row0_txt):
            for i in range(1, len(df)):
                row_list = [str(x).strip() for x in df.iloc[i]]
                if not any(row_list):
                    continue
                found_veh = None
                found_state = None
                found_terr = None
                for val in row_list:
                    v = val.strip().upper()
                    if val in veh_df["Veh No."].values and not found_veh:
                        found_veh = val
                    elif re.match(r'^[A-Z]{2}$', v) and not found_state:
                        found_state = v
                    elif re.match(r'^\d+$', v) and not found_terr:
                        found_terr = v.zfill(3)
                if found_veh:
                    classification_rows.append({
                        "Veh No.": found_veh,
                        "State_class": found_state if found_state else "",
                        "Terr_class": found_terr if found_terr else ""
                    })

    if classification_rows:
        class_df = pd.DataFrame(classification_rows)
        class_df.drop_duplicates(subset=["Veh No."], keep="last", inplace=True)
        class_df["Veh No."] = class_df["Veh No."].astype(str).str.strip()
        veh_df = pd.merge(veh_df, class_df, on="Veh No.", how="left")

        mask_no_state = (veh_df["State"] == "") & (veh_df["State_class"].notna())
        veh_df.loc[mask_no_state, "State"] = veh_df.loc[mask_no_state, "State_class"]

        mask_no_terr = (veh_df["Territory"] == "") & (veh_df["Terr_class"].notna())
        veh_df.loc[mask_no_terr, "Territory"] = veh_df.loc[mask_no_terr, "Terr_class"]

        veh_df.drop(columns=["State_class", "Terr_class"], inplace=True, errors="ignore")

    mask_unmatched = (veh_df["State"] == "") | (veh_df["Territory"] == "")
    if mask_unmatched.any():
        text_pdfminer = pdfminer_extract_text(io.BytesIO(pdf_data))
        text_clean = " ".join(text_pdfminer.split())
        m = re.search(r"\b([A-Z]{2})\s+Terr\s+(\d+)\b", text_clean)
        if m:
            fallback_state = m.group(1)
            fallback_terr = m.group(2).zfill(3)
            mask_no_state = (veh_df["State"] == "")
            mask_no_terr = (veh_df["Territory"] == "")
            veh_df.loc[mask_no_state, "State"] = veh_df.loc[mask_no_state, "State"].replace("", fallback_state)
            veh_df.loc[mask_no_terr, "Territory"] = veh_df.loc[mask_no_terr, "Territory"].replace("", fallback_terr)
    new_pairs = extract_state_territory_from_pymupdf(pdf_data)
    if new_pairs:
        for idx, (state, territory) in enumerate(new_pairs):
            if idx < len(veh_df):
                veh_df.at[idx, "State"] = state
                veh_df.at[idx, "Territory"] = territory

    return veh_df

def fallback_extract_1_5_dynamic_with_value(pdf_data):
    lines = extract_text_pymupdf(pdf_data)
    schedule_indices = []
    for i, line in enumerate(lines):
        if "schedule of covered autos you own" in line.lower():
            schedule_indices.append(i)
    if not schedule_indices:
        return pd.DataFrame(columns=["Veh No.", "Year", "Model", "VIN Number", "Value"])
    all_vehicles = []
    for start_idx in schedule_indices:
        relevant = []
        j = start_idx + 1
        while j < len(lines):
            upper_line = lines[j].upper().strip()
            if ("CLASSIFICATION" in upper_line) or ("SCHEDULE OF LOSS PAYEES" in upper_line):
                break
            relevant.append(lines[j])
            j += 1
        idx = 0
        n = len(relevant)
        while idx < n:
            line = relevant[idx].strip()
            if line.isdigit():
                veh_no = line
                year = ""
                model = ""
                vin = ""
                value_str = ""
                idx += 1
                if idx < n:
                    possible_year = relevant[idx].strip()
                    if re.match(r'^\d{4}$', possible_year):
                        year = possible_year
                        idx += 1
                while idx < n:
                    check_line = relevant[idx].replace('\n', '').strip()
                    upper_check = check_line.upper()
                    if check_line.isdigit() or "CLASSIFICATION" in upper_check or "SCHEDULE OF LOSS PAYEES" in upper_check:
                        break
                    if check_line and not check_line.startswith("$"):
                        model = check_line
                        idx += 1
                        break
                    idx += 1
                while idx < n:
                    check_line = relevant[idx].replace('\n', '').strip()
                    upper_check = check_line.upper()
                    if check_line.isdigit() or "CLASSIFICATION" in upper_check or "SCHEDULE OF LOSS PAYEES" in upper_check:
                        break
                    possible_vin = find_vin_in_text(check_line)
                    if possible_vin:
                        leftover = check_line.replace(possible_vin, "").strip()
                        if leftover:
                            model = (model + " " + leftover).strip()
                        vin = possible_vin
                        idx += 1
                        break
                    idx += 1
                while idx < n:
                    check_line = relevant[idx].replace('\n', '').strip()
                    upper_check = check_line.upper()
                    if check_line.isdigit() or "CLASSIFICATION" in upper_check or "SCHEDULE OF LOSS PAYEES" in upper_check:
                        break
                    if re.match(r'^\$?\d{1,3}(,\d{3})*(\.\d+)?$', check_line):
                        value_str = check_line
                        idx += 1
                        break
                    idx += 1
                if not value_str:
                    temp_idx = idx
                    chunk_for_this_vehicle = []
                    while temp_idx < n:
                        temp_line = relevant[temp_idx].replace('\n', '').strip()
                        temp_upper = temp_line.upper()
                        if temp_line.isdigit() or "CLASSIFICATION" in temp_upper or "SCHEDULE OF LOSS PAYEES" in temp_upper:
                            break
                        chunk_for_this_vehicle.append(temp_line)
                        temp_idx += 1
                    joined_chunk = " ".join(chunk_for_this_vehicle)
                    fallback_match = re.search(r'\d{1,3}(,\d{3})+(\.\d+)?', joined_chunk)
                    if fallback_match:
                        value_str = fallback_match.group(0)
                while idx < n:
                    skip_line = relevant[idx].strip()
                    upper_skip = skip_line.upper()
                    if skip_line.isdigit() or "CLASSIFICATION" in upper_skip or "SCHEDULE OF LOSS PAYEES" in upper_skip:
                        break
                    idx += 1
                row_dict = {
                    "Veh No.": veh_no,
                    "Year": year,
                    "Model": model.strip(),
                    "VIN Number": vin.strip(),
                    "Value": value_str.strip()
                }
                all_vehicles.append(row_dict)
            else:
                idx += 1
    fallback_df = pd.DataFrame(all_vehicles)
    if not fallback_df.empty:
        fallback_df = fallback_df[fallback_df["Veh No."].str.isdigit()].copy()
        fallback_df.reset_index(drop=True, inplace=True)
    return fallback_df

def separate_vin_in_model(veh_df):
    for i in range(len(veh_df)):
        model_val = veh_df.at[i, "Model"] or ""
        possible_vin = find_vin_in_text(model_val)
        if possible_vin:
            leftover = model_val.replace(possible_vin, "").strip()
            veh_df.at[i, "Model"] = leftover
            if not veh_df.at[i, "VIN Number"]:
                veh_df.at[i, "VIN Number"] = possible_vin
    return veh_df

def final_vin_cleanup(veh_df):
    import re
    pattern = re.compile(r"[A-HJ-NPR-Z0-9]{15,17}", re.IGNORECASE)
    for i in range(len(veh_df)):
        vin_col = veh_df.at[i, "VIN Number"] or ""
        if not looks_like_vin(vin_col):
            model_val = veh_df.at[i, "Model"] or ""
            match = pattern.search(model_val)
            if match:
                possible_vin = match.group(0).upper()
                leftover = re.sub(match.group(0), "", model_val, flags=re.IGNORECASE).strip()
                veh_df.at[i, "Model"] = leftover
                veh_df.at[i, "VIN Number"] = possible_vin
    return veh_df

def extract_premium_camelot(pdf_data):
    import tempfile, os, re
    import camelot
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_data)
        tmp.flush()
        tmp_name = tmp.name
    tables = camelot.read_pdf(tmp_name, flavor='stream', pages='all')
    os.remove(tmp_name)
    def format_premium_with_commas(num_str):
        try:
            val = float(num_str.replace("$", "").replace(",", "").strip())
            if val.is_integer():
                return f"{int(val):,}"
            else:
                formatted = f"{val:,.2f}".rstrip('0').rstrip('.')
                return formatted
        except:
            return num_str
    premium_map = {}
    for t in tables:
        df = t.df.copy()
        df = df.reset_index(drop=True)
        df.columns = pd.Series(df.columns.astype(str))
        df = df.loc[:, ~df.columns.duplicated()]
        df_text = df.to_string().lower()
        if "premium" not in df_text:
            continue
        header_idx = None
        for i2, row_vals in df.iterrows():
            row_str = " ".join(str(x).lower() for x in row_vals)
            if "no." in row_str and "premium" in row_str:
                header_idx = i2
                break
        if header_idx is None:
            continue
        df.columns = df.iloc[header_idx]
        df = df.drop(range(header_idx + 1))
        df = df.reset_index(drop=True)
        if "No." not in df.columns or "Premium" not in df.columns:
            continue
        last_col = df.columns[-1]
        for i2 in range(len(df)):
            row_vals = df.iloc[i2].tolist()
            row_str_lower = " ".join(str(x).lower() for x in row_vals).strip()
            if "schedule of loss payees" in row_str_lower:
                break
            veh_no = str(df.iloc[i2].get("No.", "")).strip()
            if not veh_no.isdigit():
                continue
            raw_prem = str(df.iloc[i2].get(last_col, "")).strip()
            money_matches = re.findall(r"\$?\s*\d[\d,\.]*", raw_prem)
            if money_matches:
                final_prem = money_matches[-1]
            else:
                row_full_str = " ".join(str(x) for x in row_vals)
                row_matches = re.findall(r"\$?\s*\d[\d,\.]*", row_full_str)
                final_prem = row_matches[-1] if row_matches else ""
            if final_prem:
                premium_map[veh_no] = format_premium_with_commas(final_prem)
    return premium_map

def extract_table3_camelot(pdf_data):
    import re
    import tempfile, os
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_data)
        tmp.flush()
        tmp_name = tmp.name
    tables = camelot.read_pdf(tmp_name, flavor='stream', pages='all')
    os.remove(tmp_name)
    all_vehicles = []
    for t in tables:
        df = t.df.copy()
        df = df.reset_index(drop=True)
        df.columns = pd.Series(df.columns.astype(str))
        df = df.loc[:, ~df.columns.duplicated()]
        txt = df.to_string().lower()
        if ("schedule of covered autos you own" in txt and "no." in txt and "year" in txt and "vin" in txt):
            header_idx = None
            for i, row in df.iterrows():
                row_text = " ".join(str(x).lower() for x in row)
                if ("no." in row_text and "year" in row_text and "vin" in row_text):
                    header_idx = i
                    break
            if header_idx is not None:
                df.columns = df.iloc[header_idx]
                df = df.drop(range(header_idx + 1))
                df = df.reset_index(drop=True)
                needed = ["No.", "Year", "Model", "VIN Number"]
                if all(x in df.columns for x in needed):
                    merged_rows = []
                    prev = None
                    for idx2 in range(len(df)):
                        row_vals = df.iloc[idx2].tolist()
                        row_vals = [str(x).strip() for x in row_vals]
                        row_lower = " ".join(row_vals).lower()
                        if ("classification" in row_lower) or ("schedule of loss payees" in row_lower):
                            break
                        row_dict = {
                            "Veh No.": row_vals[0] if len(row_vals) > 0 else "",
                            "Year": row_vals[1] if len(row_vals) > 1 else "",
                            "Model": row_vals[2] if len(row_vals) > 2 else "",
                            "VIN Number": row_vals[3] if len(row_vals) > 3 else ""
                        }
                        row_dict["Value"] = ""
                        if len(row_vals) >= 7:
                            col4 = row_vals[4]
                            col5 = row_vals[5]
                            col6 = row_vals[6]
                            if col4.startswith("$") and re.search(r"\d", col5):
                                row_dict["Value"] = col4 + col5
                            else:
                                if col5.startswith("$") and re.search(r"\d", col5):
                                    row_dict["Value"] = col5
                                elif col6.startswith("$") and re.search(r"\d", col6):
                                    row_dict["Value"] = col6
                        elif len(row_vals) == 6:
                            col4 = row_vals[4]
                            col5 = row_vals[5]
                            if col4.startswith("$") and re.search(r"\d", col5):
                                row_dict["Value"] = col4 + col5
                            else:
                                if col5.startswith("$"):
                                    row_dict["Value"] = col5
                        if not row_dict["Veh No."] and prev is not None:
                            if row_dict["Year"] and not prev["Year"]:
                                prev["Year"] = row_dict["Year"]
                            if row_dict["Model"] and not prev["Model"]:
                                prev["Model"] = row_dict["Model"]
                            if row_dict["VIN Number"]:
                                if (not looks_like_vin(prev["VIN Number"])) and looks_like_vin(row_dict["VIN Number"]):
                                    prev["VIN Number"] = row_dict["VIN Number"]
                            if row_dict["Value"] and not prev["Value"]:
                                prev["Value"] = row_dict["Value"]
                        else:
                            if prev is not None:
                                merged_rows.append(prev)
                            prev = row_dict
                    if prev is not None:
                        merged_rows.append(prev)
                    all_vehicles.extend(merged_rows)
    if not all_vehicles:
        veh_df = pd.DataFrame(columns=["Veh No.","Year","Model","VIN Number","State","Territory","Value"])
    else:
        veh_df = pd.DataFrame(all_vehicles)
        veh_df = veh_df[veh_df["Veh No."].astype(str).str.strip() != ""].copy()
        veh_df.reset_index(drop=True, inplace=True)
        if "Value" not in veh_df.columns:
            veh_df["Value"] = ""
    fallback_df = fallback_extract_1_5_dynamic_with_value(pdf_data)
    if not fallback_df.empty:
        combined = pd.concat([veh_df, fallback_df], ignore_index=True)
        combined.drop_duplicates(subset=["Veh No."], keep="last", inplace=True)
        combined.reset_index(drop=True, inplace=True)
        veh_df = combined
    try:
        veh_df["Veh No."] = veh_df["Veh No."].astype(int)
        veh_df.sort_values(by="Veh No.", inplace=True)
        veh_df.reset_index(drop=True, inplace=True)
    except:
        pass
    veh_df = separate_vin_in_model(veh_df)
    veh_df = final_vin_cleanup(veh_df)
    if "State" not in veh_df.columns:
        veh_df["State"] = ""
    if "Territory" not in veh_df.columns:
        veh_df["Territory"] = ""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_data)
        tmp.flush()
        tmp_name = tmp.name
    tables_for_merge = camelot.read_pdf(tmp_name, flavor='stream', pages='all')
    os.remove(tmp_name)
    veh_df = merge_classification_and_territory(veh_df, tables_for_merge, pdf_data)
    def extract_premium_pypdf_for_table3(pdf_bytes):
        from pypdf import PdfReader
        import io
        import re
        reader = PdfReader(io.BytesIO(pdf_bytes))
        text = ""
        for page in reader.pages:
            page_text = page.extract_text() or ""
            text += page_text + "\n"
        lines = text.splitlines()
        start_idx = None
        for i, line in enumerate(lines):
            if "PHYSICAL DAMAGE COVERAGE" in line.upper():
                start_idx = i
                break
        if start_idx is None:
            return {}
        relevant = []
        for i in range(start_idx + 1, len(lines)):
            upper_line = lines[i].upper()
            if "SCHEDULE OF LOSS PAYEES" in upper_line:
                break
            relevant.append(lines[i])
        unified = []
        buffer_line = ""
        skip_keywords = re.compile(r'(original cost|newstated|year model vin number)', re.IGNORECASE)
        for line in relevant:
            line_str = line.strip()
            if re.match(r'^\d+\b', line_str):
                if buffer_line.strip():
                    unified.append(buffer_line)
                buffer_line = line_str
            else:
                if skip_keywords.search(line_str):
                    if buffer_line.strip():
                        unified.append(buffer_line)
                    buffer_line = ""
                    continue
                else:
                    buffer_line += " " + line_str
        if buffer_line.strip():
            unified.append(buffer_line)
        results_dict = {}
        for line in unified:
            line = line.strip()
            m = re.match(r'^(\d+)\b', line)
            if m:
                veh_no = m.group(1)
                money_matches = re.findall(r'\$\s*\d[\d,\.]*', line)
                if money_matches:
                    last_val = money_matches[-1]
                    results_dict[veh_no] = last_val
        return results_dict
    premium_mapping_pypdf = extract_premium_pypdf_for_table3(pdf_data)
    veh_df["Premium"] = ""
    for idx, row in veh_df.iterrows():
        veh_no = str(row["Veh No."]).strip()
        if veh_no in premium_mapping_pypdf:
            veh_df.at[idx, "Premium"] = premium_mapping_pypdf[veh_no]
    veh_df["Value"] = veh_df["Value"].astype(str)
    veh_df["Premium"] = veh_df["Premium"].astype(str)
    for idx in range(len(veh_df)):
        val_str = veh_df.at[idx, "Value"].strip()
        if val_str and not val_str.startswith("$"):
            veh_df.at[idx, "Value"] = f"${val_str}"
        prem_str = veh_df.at[idx, "Premium"].strip()
        if prem_str and not prem_str.startswith("$"):
            veh_df.at[idx, "Premium"] = f"${prem_str}"
    return veh_df

def extract_loss_payees(pdf_data):
    lines = extract_text_pymupdf(pdf_data)
    schedule_autos_idx = None
    for i, line in enumerate(lines):
        if "schedule of covered autos you own" in line.lower():
            schedule_autos_idx = i
            break
    if schedule_autos_idx is None:
        return []
    schedule_loss_idx = None
    for i in range(schedule_autos_idx + 1, len(lines)):
        if "schedule of loss payees" in lines[i].lower():
            schedule_loss_idx = i
            break
    if schedule_loss_idx is None:
        return []
    end_idx = len(lines)
    for i in range(schedule_loss_idx + 1, len(lines)):
        if "schedule of hired or borrowed covered auto coverage and premiums" in lines[i].lower():
            end_idx = i
            break
    relevant = lines[schedule_loss_idx + 1 : end_idx]
    payees = []
    found_veh_no = False
    for line in relevant:
        line_str = line.strip()
        if not line_str:
            continue
        match = re.match(r'^(\d+)\s+(.*)', line_str)
        if match:
            veh = match.group(1).strip()
            payee = match.group(2).strip()
            payees.append({"Veh No.": veh, "Loss Payee": payee})
            found_veh_no = True
    if not found_veh_no:
        payees.append({"Veh No.": "None", "Loss Payee": "None"})
    return payees

def extract_cost_of_hire_used_pdfplumber(pdf_data):
    import tempfile
    import re
    coverage_data = {"Primary Coverage": {"State": "-", "Premium": "-"},
                     "Excess Coverage":  {"State": "-", "Premium": "-"}}
    lines = []
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_data)
        tmp.flush()
        tmp_name = tmp.name
    try:
        with pdfplumber.open(tmp_name) as plumber_pdf:
            for page in plumber_pdf.pages:
                text = page.extract_text() or ""
                for ln in text.splitlines():
                    lines.append(ln.strip())
    except:
        pass
    os.remove(tmp_name)
    start_idx = None
    for i in range(len(lines)-1):
        if ("liability coverage - cost of hire rating basis for autos used in your motor carrier operations" in lines[i].lower() and
            "(other than mobile or farm equipment)" in lines[i+1].lower()):
            start_idx = i
            break
    if start_idx is None:
        return pd.DataFrame(columns=["Coverage","State","Premium"])
    stop_phrases = ["total premiums:",
                    "liability coverage - cost of hire rating basis for autos not used in your motor carrier operations"]
    i = start_idx + 2
    n = len(lines)
    while i < n:
        low_line = lines[i].lower()
        if any(sp in low_line for sp in stop_phrases):
            break
        for coverage_key in ["Primary Coverage", "Excess Coverage"]:
            if coverage_key.lower() in low_line:
                state_match = re.search(r'\b[A-Z]{2}\b', lines[i])
                if not state_match and i+1 < n:
                    next_line = lines[i+1].lower()
                    if not any(sp in next_line for sp in stop_phrases):
                        state_match = re.search(r'\b[A-Z]{2}\b', lines[i+1])
                if state_match:
                    coverage_data[coverage_key]["State"] = state_match.group(0)
                numeric_match = re.search(r'\$?\d[\d,\.]*', low_line)
                if not numeric_match and i+1 < n:
                    nxt_low = lines[i+1].lower()
                    if not any(sp in nxt_low for sp in stop_phrases):
                        numeric_match = re.search(r'\$?\d[\d,\.]*', nxt_low)
                if numeric_match:
                    coverage_data[coverage_key]["Premium"] = numeric_match.group(0)
        i += 1
    rows = []
    for key in ["Primary Coverage", "Excess Coverage"]:
        rows.append({"Coverage": key,
                     "State": coverage_data[key]["State"],
                     "Premium": coverage_data[key]["Premium"]})
    return pd.DataFrame(rows, columns=["Coverage","State","Premium"])

def extract_cost_of_hire_not_used_pdfplumber(pdf_data):
    import tempfile, re
    coverage_data = {"Primary Coverage": {"State": "-", "Premium": "-"},
                     "Excess Coverage":  {"State": "-", "Premium": "-"}}
    lines = []
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_data)
        tmp.flush()
        tmp_name = tmp.name
    try:
        with pdfplumber.open(tmp_name) as plumber_pdf:
            for page in plumber_pdf.pages:
                text = page.extract_text() or ""
                for ln in text.splitlines():
                    lines.append(ln.strip())
    except:
        pass
    os.remove(tmp_name)
    for_not_used = "liability coverage - cost of hire rating basis for autos not used in your motor carrier operations"
    other_line = "(other than mobile or farm equipment)"
    start_idx = None
    for i in range(len(lines)-1):
        if (for_not_used in lines[i].lower() and other_line in lines[i+1].lower()):
            start_idx = i
            break
    if start_idx is None:
        return pd.DataFrame(columns=["Coverage","State","Premium"])
    stop_phrases = ['for"auto"notusedin',
                    'for "autos" used in your motor carrier operations, cost of hire means:',
                    'total premiums:']
    i = start_idx + 2
    n = len(lines)
    while i < n:
        low_line = lines[i].lower()
        if any(sp in low_line for sp in stop_phrases):
            break
        for coverage_key in ["Primary Coverage", "Excess Coverage"]:
            if coverage_key.lower() in low_line:
                state_match = re.search(r'\b[A-Z]{2}\b', lines[i])
                if not state_match and i+1 < n:
                    nxt = lines[i+1].lower()
                    if not any(sp in nxt for sp in stop_phrases):
                        state_match = re.search(r'\b[A-Z]{2}\b', nxt)
                if state_match:
                    coverage_data[coverage_key]["State"] = state_match.group(0)
                numeric_match = re.search(r'\$?\d[\d,\.]*', low_line)
                if not numeric_match and i+1 < n:
                    nxt_low = lines[i+1].lower()
                    if not any(sp in nxt_low for sp in stop_phrases):
                        numeric_match = re.search(r'\$?\d[\d,\.]*', nxt_low)
                if numeric_match:
                    coverage_data[coverage_key]["Premium"] = numeric_match.group(0)
        i += 1
    rows = []
    for key in ["Primary Coverage", "Excess Coverage"]:
        rows.append({"Coverage": key,
                     "State": coverage_data[key]["State"],
                     "Premium": coverage_data[key]["Premium"]})
    return pd.DataFrame(rows, columns=["Coverage","State","Premium"])

def extract_non_ownership_liability_pymupdf(pdf_data):
    import fitz, re
    forced_rows = [
        {"Business": "Other Than Auto Service", "Basis": "Number Of Employees"},
        {"Business": "Operations, Partnerships Or LLCs", "Basis": "Number Of Volunteers"},
        {"Business": "Auto Service Operations", "Basis": "Number Of Employees Whose Principal Duty Involves The Operation Of Autos"},
        {"Business": "", "Basis": "Number Of Volunteers"},
        {"Business": "", "Basis": "Number Of Partners (Active And Inactive) Or LLC Members"},
        {"Business": "Partnerships Or LLCs", "Basis": "Number Of Employees"},
        {"Business": "", "Basis": "Number Of Volunteers"},
        {"Business": "", "Basis": "Number Of Partners (Active And Inactive) Or LLC Members"}
    ]
    doc = fitz.open(stream=pdf_data, filetype="pdf")
    all_lines = []
    for page in doc:
        text = page.get_text() or ""
        for ln in text.splitlines():
            all_lines.append(ln.strip())
    doc.close()
    start_idx = None
    for i, line in enumerate(all_lines):
        if "schedule for non-ownership liability" in line.lower():
            start_idx = i
            break
    # No Non-Ownership Liability table found -> return empty
    if start_idx is None:
        return pd.DataFrame(columns=[
            "Named Insured's Business",
            "Rating Basis",
            "Number",
            "Premium"
        ])
    stop_idx = None
    for i in range(start_idx+1, len(all_lines)):
        if "additional coverages" in all_lines[i].lower():
            stop_idx = i
            break
    if stop_idx is None:
        stop_idx = len(all_lines)
    relevant = all_lines[start_idx:stop_idx]
    unified = []
    skip_next = False
    for i in range(len(relevant)):
        if skip_next:
            skip_next = False
            continue
        line = relevant[i]
        if i+1 < len(relevant):
            next_line = relevant[i+1]
            if line.lower().endswith("principal") and next_line.lower().startswith("duty involves"):
                unified.append(line + " " + next_line)
                skip_next = True
                continue
        unified.append(line)
    search_start = 0
    final_data = []
    for forced_row in forced_rows:
        basis_text = forced_row["Basis"].lower()
        found_number = ""
        found_premium = ""
        matched_line_idx = None
        for idx in range(search_start, len(unified)):
            if basis_text in unified[idx].lower():
                matched_line_idx = idx
                search_start = matched_line_idx + 1
                break
        if matched_line_idx is not None:
            for offset in [0,1,2]:
                check_idx = matched_line_idx + offset
                if check_idx >= len(unified):
                    break
                check_line = unified[check_idx]
                if not found_number and '$' not in check_line:
                    nm = re.search(r'\b(\d+)\b', check_line)
                    if nm:
                        found_number = nm.group(1)
                if not found_premium:
                    pm = re.search(r'\$\d[\d,\.]*', check_line)
                    if pm:
                        found_premium = pm.group(0)
                if found_number and found_premium:
                    break
        final_data.append({
            "Named Insured's Business": forced_row["Business"],
            "Rating Basis": forced_row["Basis"],
            "Number": found_number,
            "Premium": found_premium
        })
    # Build DataFrame and only return if any Premium exists
    df_non_own = pd.DataFrame(final_data, columns=["Named Insured's Business","Rating Basis","Number","Premium"])
    if df_non_own['Premium'].fillna('').astype(str).str.strip().eq('').all():
        return pd.DataFrame(columns=df_non_own.columns)
    return df_non_own

def extract_additional_coverages_pymupdf(pdf_data):
    import fitz
    import re
    import pandas as pd

    doc = fitz.open(stream=pdf_data, filetype="pdf")
    all_lines = []
    for page in doc:
        text = page.get_text() or ""
        for ln in text.splitlines():
            all_lines.append(ln.strip())
    doc.close()

    # Locate "additional coverages"
    addl_idx = None
    for i, line in enumerate(all_lines):
        if "additional coverages" in line.lower():
            addl_idx = i
            break
    if addl_idx is None:
        return pd.DataFrame(columns=["Coverage", "Limit", "Deductible", "Premium"])

    # After "additional coverages", locate "product wide coverages"
    prod_idx = None
    for i in range(addl_idx + 1, len(all_lines)):
        if "product wide coverages" in all_lines[i].lower():
            prod_idx = i
            break
    if prod_idx is None:
        return pd.DataFrame(columns=["Coverage", "Limit", "Deductible", "Premium"])

    # Stop when we hit specific stop phrases such as "location coverages" or "vehicle coverages"
    stop_phrases = [
        "location coverages",
        "vehicle coverages",
        "commercial liability umbrella quote proposal",
        "commercial inland marine quote proposal",
    ]
    stop_idx = None
    for i in range(prod_idx + 1, len(all_lines)):
        lower_line = all_lines[i].lower()
        if any(sp in lower_line for sp in stop_phrases):
            stop_idx = i
            break
    if stop_idx is None:
        stop_idx = len(all_lines)

    # Extract block of lines between "product wide coverages" and stopping phrase
    block = all_lines[prod_idx + 1 : stop_idx]

    # Filter out header rows and blank lines (like "Coverage", "Limit", etc.)
    header_words = {"coverage", "limit", "deductible", "premium", ""}
    filtered = []
    for ln in block:
        low_ln = ln.lower().strip()
        if low_ln in header_words:
            continue
        if not ln.strip():
            continue
        filtered.append(ln)

    # Parse the filtered lines into 4 columns: Coverage, Limit, Deductible, Premium.
    results = []
    i = 0
    n = len(filtered)

    while i < n:
        coverage = filtered[i]
        limit = ""
        deductible = ""
        premium = ""
        i += 1

        # Look at the next few lines (up to 4) for potential values
        reads = 0
        while reads < 4 and i < n:
            line = filtered[i]
            numeric = re.sub(r"[^\d]", "", line)
            if line.startswith("$"):
                if len(numeric) >= 5:
                    limit = line
                else:
                    premium = line
            elif re.match(r"^\$?\d[\d,\.]*$", line):
                if len(numeric) >= 5:
                    limit = line
                else:
                    if premium:
                        deductible = line
                    else:
                        premium = line
            i += 1
            reads += 1

        results.append({
            "Coverage": coverage.strip(),
            "Limit": limit.strip(),
            "Deductible": deductible.strip(),
            "Premium": premium.strip()
        })

    # PATCH: If Premium is empty and Deductible starts with a dollar sign, move its value to Premium.
    for row in results:
        if not row["Premium"] and row["Deductible"].startswith("$"):
            row["Premium"] = row["Deductible"]
            row["Deductible"] = ""

    df = pd.DataFrame(results, columns=["Coverage", "Limit", "Deductible", "Premium"])
    return df

def extract_vehicle_coverages_pymupdf(pdf_data):
    import fitz, re
    doc = fitz.open(stream=pdf_data, filetype="pdf")
    all_lines = []
    for page in doc:
        text = page.get_text() or ""
        for ln in text.splitlines():
            all_lines.append(ln.strip())
    doc.close()
    start_idx = None
    for i, line in enumerate(all_lines):
        if "vehicle coverages" in line.lower():
            start_idx = i
            break
    # No Non-Ownership Liability table found -> return empty
    if start_idx is None:
        return pd.DataFrame(columns=[
            "Named Insured's Business",
            "Rating Basis",
            "Number",
            "Premium"
        ])
    data_start = start_idx + 6
    stop_phrases = [
        "location coverages",
        "commercial liability umbrella quote proposal",
        "commercial inland marine quote proposal",
        "additional coverages",
        "rating company:"
    ]
    stop_idx = None
    for j in range(data_start, len(all_lines)):
        line_low = all_lines[j].lower()
        if any(sp in line_low for sp in stop_phrases):
            stop_idx = j
            break
    if stop_idx is None:
        stop_idx = len(all_lines)
    relevant = all_lines[data_start:stop_idx]
    if not relevant:
        return pd.DataFrame(
            [{"Veh#":"-","Coverage":"-","Limit":"-","Deductible":"-","Premium":"-"}],
            columns=["Veh#","Coverage","Limit","Deductible","Premium"]
        )
    rows = []
    for line in relevant:
        parts = line.split()
        while len(parts) < 5:
            parts.append("-")
        parts = parts[:5]
        rows.append({"Veh#": parts[0], "Coverage": parts[1], "Limit": parts[2],
                     "Deductible": parts[3], "Premium": parts[4]})
    return pd.DataFrame(rows, columns=["Veh#","Coverage","Limit","Deductible","Premium"])

def extract_location_coverages_pymupdf(pdf_data):
    """
    Extract 'Location Coverages' table, stopping at stop phrases or 'PROPOSAL 01 00'.
    """
    import fitz, re
    # pd is already imported at module level
    doc = fitz.open(stream=pdf_data, filetype="pdf")
    all_lines = []
    for page in doc:
        text = page.get_text() or ""
        for ln in text.splitlines():
            all_lines.append(ln.strip())
    doc.close()
    # Find the 'Location Coverages' header
    start_idx = None
    for idx, line in enumerate(all_lines):
        if "location coverages" in line.lower():
            start_idx = idx
            break
    if start_idx is None:
        return pd.DataFrame(columns=["Location","Coverage","Limit","Deductible","Premium"])
    # Define where data starts
    data_start = start_idx + 1
    stop_phrases = [
        "commercial liability umbrella quote proposal",
        "commercial inland marine quote proposal",
        "vehicle coverages",
        "additional coverages"
    ]
    # Also stop at 'proposal 01 00'
    stop_idx = None
    for j in range(data_start, len(all_lines)):
        low_line = all_lines[j].lower()
        if "proposal 01 00" in low_line or any(sp in low_line for sp in stop_phrases):
            stop_idx = j
            break
    if stop_idx is None:
        stop_idx = len(all_lines)
    relevant = all_lines[data_start:stop_idx]
    # Skip subheaders and blanks
    subheaders = {"location","coverage","limit","deductible","premium",""}
    def skip_subheaders_and_blank(idx):
        while idx < len(relevant) and relevant[idx].strip().lower() in subheaders:
            idx += 1
        return idx
    idx = skip_subheaders_and_blank(0)
    rows = []
    current_location = ""
    # Parse rows
    while idx < len(relevant):
        line = relevant[idx].strip()
        low_line = line.lower()
        # Detect explicit 'Location:' lines
        if low_line.startswith("location:"):
            current_location = line.split(":",1)[1].strip()
            idx += 1
            idx = skip_subheaders_and_blank(idx)
            continue
        # Ensure enough lines remain
        if idx + 3 >= len(relevant):
            break
        coverage = line
        idx += 1
        idx = skip_subheaders_and_blank(idx)
        limit = relevant[idx].strip() if idx < len(relevant) else "-"
        idx += 1
        idx = skip_subheaders_and_blank(idx)
        deductible = relevant[idx].strip() if idx < len(relevant) else "-"
        idx += 1
        idx = skip_subheaders_and_blank(idx)
        # Sometimes a lone '$' precedes premium
        if idx < len(relevant) and relevant[idx].strip() == "$":
            idx += 1
        idx = skip_subheaders_and_blank(idx)
        premium = relevant[idx].strip() if idx < len(relevant) else "-"
        idx += 1
        rows.append({
            "Location": current_location or "-",
            "Coverage": coverage or "-",
            "Limit": limit or "-",
            "Deductible": deductible or "-",
            "Premium": premium or "-"
        })
        idx = skip_subheaders_and_blank(idx)
    # If no rows found, return a placeholder row
    if not rows:
        return pd.DataFrame([{"Location":"-","Coverage":"-","Limit":"-","Deductible":"-","Premium":"-"}])
    return pd.DataFrame(rows, columns=["Location","Coverage","Limit","Deductible","Premium"])

    subheaders = {"location","coverage","limit","deductible","premium",""}

    def skip_subheaders_and_blank(idx):
        while idx < len(relevant):
            test_line = relevant[idx].strip().lower()
            if test_line in subheaders:
                idx += 1
            else:
                break
        return idx

    idx = skip_subheaders_and_blank(0)
    rows = []
    current_location = ""
    found_any_coverage = False
    n = len(relevant)

    while idx < n:
        line = relevant[idx].strip()
        low_line = line.lower()

        if low_line.startswith("location:"):
            current_location = line.split(":",1)[1].strip()
            idx += 1
            idx = skip_subheaders_and_blank(idx)
            continue

        if low_line in subheaders:
            idx += 1
            continue

        if idx + 4 >= n:
            break

        coverage_line = line
        idx += 1

        idx = skip_subheaders_and_blank(idx)
        limit_line = relevant[idx].strip() if idx < n else "-"
        idx += 1

        idx = skip_subheaders_and_blank(idx)
        ded_line = relevant[idx].strip() if idx < n else "-"
        idx += 1

        idx = skip_subheaders_and_blank(idx)
        if idx < n and relevant[idx].strip() == "$":
            idx += 1
        idx = skip_subheaders_and_blank(idx)

        prem_line = relevant[idx].strip() if idx < n else "-"
        idx += 1

        if re.match(r'^\d+(\.\d+)?$', prem_line.replace(",","")):
            try:
                val = float(prem_line.replace(",",""))
                prem_line = f"${int(val):,}"
            except:
                pass

        loc_for_this = current_location if current_location else ""
        rows.append({
            "Location":   loc_for_this or "-",
            "Coverage":   coverage_line or "-",
            "Limit":      limit_line or "-",
            "Deductible": ded_line or "-",
            "Premium":    prem_line or "-"
        })
        found_any_coverage = True

        idx = skip_subheaders_and_blank(idx)

    if not found_any_coverage:
        return pd.DataFrame(
            [{"Location":"-","Coverage":"-","Limit":"-","Deductible":"-","Premium":"-"}],
            columns=["Location","Coverage","Limit","Deductible","Premium"]
        )

    return pd.DataFrame(rows, columns=["Location","Coverage","Limit","Deductible","Premium"])

##############################################################################
# POLICY FORMS EXTRACTION
##############################################################################
def extract_text_pdfplumber_custom(pdf_bytes: bytes) -> str:
    try:
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            extracted_text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
        return extracted_text
    except Exception as e:
        st.error(f"Error with pdfplumber: {e}")
        return ""

def extract_text_pymupdf_custom(pdf_bytes: bytes) -> str:
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        extracted_text = "\n".join(page.get_text() for page in doc)
        doc.close()
        return extracted_text
    except Exception as e:
        st.error(f"Error with PyMuPDF: {e}")
        return ""

def parse_line_into_columns(line: str) -> list:
    tokens = line.split()
    number_parts = []
    edition = ""
    description_parts = []
    edition_pattern = re.compile(r"^\d{2}-\d{4}$")
    found_edition = False

    for token in tokens:
        if not found_edition and edition_pattern.match(token):
            edition = token
            found_edition = True
        else:
            if not found_edition:
                number_parts.append(token)
            else:
                description_parts.append(token)

    number = " ".join(number_parts).strip()
    description = " ".join(description_parts).strip()
    return [number, edition, description]




def parse_policy_forms(text: str) -> dict:
    """
    Extract Policy Forms sections for all coverage_titles, parsing table rows until
    narrative markers appear or duplicate rows indicate we've looped. Returns a dict of section name -> rows.
    """
    import re
    lines = text.splitlines()
    sections = {}
    coverage_titles = ["Commercial Auto Coverage Part", "Commercial Auto"]
    # Narrative markers to stop parsing rows
    narrative_markers = ["policy number:", "applicant", "effective date:", "commercial automobile cl pn", "all commercial inland marine coverages"]
    # Pattern for a new table row: code + space + edition (dd-yyyy)
    row_pattern = re.compile(r'^[A-Z0-9 ]+\s\d{2}-\d{4}\b')

    for title in coverage_titles:
        # Find the title line
        idx = next((i for i, l in enumerate(lines) if l.strip() == title), None)
        if idx is None:
            continue
        # Find header row with Number, Edition, Description
        header_idx = next((j for j in range(idx+1, len(lines))
                           if all(k in lines[j].lower() for k in ["number", "edition", "description"])), None)
        if header_idx is None:
            continue
        rows = []
        seen = set()
        # Parse rows after header
        for line in lines[header_idx + 1:]:
            txt = line.strip()
            low = txt.lower()
            # Stop if narrative markers are encountered
            if any(marker in low for marker in narrative_markers):
                break
            # If line matches a new row
            if row_pattern.match(txt):
                num, edt, desc = parse_line_into_columns(txt)
                # Break on duplicated row tuple
                if (num, edt) in seen:
                    break
                seen.add((num, edt))
                rows.append([num, edt, clean_description(desc)])
            else:
                # Continuation of previous description
                if rows and txt:
                    prev_desc = rows[-1][2]
                    rows[-1][2] = clean_description(prev_desc + " " + txt)
        if rows:
            sections[title] = rows
    return sections
##############################################################################
# MAIN FUNCTION
##############################################################################
def main():
    st.set_page_config(page_title="Three Tables from a PDF", layout="wide", initial_sidebar_state="collapsed")
    st.title("Extracting Three Tables from a PDF")
    st.markdown(custom_css, unsafe_allow_html=True)
    
    st.header("Extracted Tables")
    
    pdf_file = st.file_uploader("Upload a PDF", type=["pdf"])
    if pdf_file is not None:
        pdf_data = pdf_file.read()
        
        # TABLE 1
        table1 = extract_table1_pypdf(pdf_data)
        # TABLE 2
        table2 = extract_table2_pymupdf(pdf_data)
        # TABLE 3
        table3 = extract_table3_camelot(pdf_data)
        
        # COVERAGE SUMMARY
        coverage_summary = table3.copy(deep=True)
        for col in ["Value", "State", "Territory", "Premium"]:
            if col in coverage_summary.columns:
                coverage_summary.drop(columns=col, inplace=True)
        if "VIN Number" in coverage_summary.columns:
            ix = coverage_summary.columns.get_loc("VIN Number")
            coverage_summary.insert(ix+1, "Liability", "")
        universal_liability = ""
        for row in table2:
            coverage_str = row.get("Coverages, Limits & Deductibles", "").lower()
            if "liability" in coverage_str:
                universal_liability = row.get("Limits", "")
                break
        coverage_summary["Liability"] = universal_liability
        
        # PDFplumber-based extraction for PIP, Med Pay, UM, UIM
        premium_details = extract_premium_pdfplumber_for_table4(pdf_data)
        for col in ["PIP", "Med Pay", "UM", "UIM"]:
            coverage_summary[col] = ""
        for idx, row in coverage_summary.iterrows():
            veh_no = str(row["Veh No."]).strip()
            if veh_no in premium_details:
                details = premium_details[veh_no]
                coverage_summary.at[idx, "PIP"] = details.get("PIP", "N")
                coverage_summary.at[idx, "Med Pay"] = details.get("Med Pay", "N")
                coverage_summary.at[idx, "UM"] = details.get("UM", "N")
                coverage_summary.at[idx, "UIM"] = details.get("UIM", "N")
            else:
                coverage_summary.at[idx, "PIP"] = "N"
                coverage_summary.at[idx, "Med Pay"] = "N"
                coverage_summary.at[idx, "UM"] = "N"
                coverage_summary.at[idx, "UIM"] = "N"
        
        deductibles = extract_deductibles_pypdf(pdf_data)
        coverage_summary["Comp\nDeductible"] = ""
        coverage_summary["Collision\nDeductible"] = ""
        for idx, row in coverage_summary.iterrows():
            veh_no = str(row["Veh No."]).strip()
            if veh_no in deductibles:
                coverage_summary.at[idx, "Comp\nDeductible"] = deductibles[veh_no].get("Comp Deductible", "")
                coverage_summary.at[idx, "Collision\nDeductible"] = deductibles[veh_no].get("Collision Deductible", "")
        
        # TABLE 5: Loss Payees
        loss_payees = extract_loss_payees(pdf_data)
        # TABLE 6: Cost of Hire (Used)
        cost_of_hire_used_df = extract_cost_of_hire_used_pdfplumber(pdf_data)
        # TABLE 7: Cost of Hire (NOT Used)
        cost_of_hire_not_used_df = extract_cost_of_hire_not_used_pdfplumber(pdf_data)
        # TABLE 8: Non-Ownership Liability
        non_ownership_df = extract_non_ownership_liability_pymupdf(pdf_data)
        # TABLE 9: Additional Coverages (updated)
        additional_coverages_df = extract_additional_coverages_pymupdf(pdf_data)
        # TABLE 10: Vehicle Coverages
        vehicle_coverages_df = extract_vehicle_coverages_pymupdf(pdf_data)
        # TABLE 11: Location Coverages
        location_coverages_df = extract_location_coverages_pymupdf(pdf_data)
        
        # Policy Forms (default to pdfplumber; sidebar removed)
        policy_text = extract_text_pdfplumber_custom(pdf_data)
        
        # ------------------------------
        # Display each table as editable HTML
        # ------------------------------
        st.subheader("Commercial Auto Coverages Premium")
        if table1:
            df1 = pd.DataFrame(table1)
            if not df1.empty:
                st.markdown(make_table_cells_editable(df1.to_html(index=False)), unsafe_allow_html=True)
            else:
                st.write("No data found for Commercial Auto Coverages Premium.")
        else:
            st.write("No data found for Commercial Auto Coverages Premium.")
        
        st.subheader("Schedule of Coverages and Covered Autos")
        if table2:
            df2 = pd.DataFrame(table2)
            if not df2.empty:
                st.markdown(make_table_cells_editable(df2.to_html(index=False)), unsafe_allow_html=True)
            else:
                st.write("No data found for Schedule of Coverages and Covered Autos.")
        else:
            st.write("No data found for Schedule of Coverages and Covered Autos.")
        
        st.subheader("Schedule of Covered Autos")
        if not table3.empty:
            st.markdown(make_table_cells_editable(table3.to_html(index=False)), unsafe_allow_html=True)
        else:
            st.write("No data found for Schedule of Covered Autos.")
        
        st.subheader("Coverage Summary")
        if not coverage_summary.empty:
            st.markdown(make_table_cells_editable(coverage_summary.to_html(index=False)), unsafe_allow_html=True)
        else:
            st.write("No data found for Coverage Summary.")
        
        st.subheader("Loss Payees")
        if loss_payees:
            df5 = pd.DataFrame(loss_payees)
            if not df5.empty:
                st.markdown(make_table_cells_editable(df5.to_html(index=False)), unsafe_allow_html=True)
            else:
                st.write("No data found for Loss Payees.")
        else:
            st.write("No data found for Loss Payees.")
        
        st.subheader("Cost of Hire (Used)")
        if not cost_of_hire_used_df.empty:
            st.markdown(make_table_cells_editable(cost_of_hire_used_df.to_html(index=False)), unsafe_allow_html=True)
        else:
            st.write("No data found for Cost of Hire (Used).")
        
        st.subheader("Cost of Hire (NOT Used)")
        if not cost_of_hire_not_used_df.empty:
            st.markdown(make_table_cells_editable(cost_of_hire_not_used_df.to_html(index=False)), unsafe_allow_html=True)
        else:
            st.write("No data found for Cost of Hire (NOT Used).")
        
        st.subheader("Non-Ownership Liability")
        if not non_ownership_df.empty:
            st.markdown(make_table_cells_editable(non_ownership_df.to_html(index=False)), unsafe_allow_html=True)
        else:
            st.write("No data found for Non-Ownership Liability.")
        
        st.subheader("Additional Coverages")
        if not additional_coverages_df.empty:
            st.markdown(make_table_cells_editable(additional_coverages_df.to_html(index=False)), unsafe_allow_html=True)
        else:
            st.write("No data found for Additional Coverages.")
        
        st.subheader("Vehicle Coverages")
        if not vehicle_coverages_df.empty:
            st.markdown(make_table_cells_editable(vehicle_coverages_df.to_html(index=False)), unsafe_allow_html=True)
        else:
            st.write("No data found for Vehicle Coverages.")
        
        st.subheader("Location Coverages")
        if not location_coverages_df.empty:
            st.markdown(make_table_cells_editable(location_coverages_df.to_html(index=False)), unsafe_allow_html=True)
        else:
            st.write("No data found for Location Coverages.")
        
        st.subheader("Policy Forms")
        if not policy_text.strip():
            st.error("No text could be extracted for Policy Forms.")
        else:
            sections = parse_policy_forms(policy_text)
            if not sections:
                st.info("No Policy Forms sections found.")
            else:
                for coverage_title, rows in sections.items():
                    st.subheader(coverage_title)
                    if rows:
                        df_policy = pd.DataFrame(rows, columns=["Number", "Edition", "Description"])
                        st.markdown(make_table_cells_editable(df_policy.to_html(index=False)), unsafe_allow_html=True)
                    else:
                        st.write("(No rows found under this coverage type.)")
    else:
        st.info("Please upload a PDF file.")

if __name__ == "__main__":
    main()