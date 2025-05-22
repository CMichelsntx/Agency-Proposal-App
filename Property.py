import pdfplumber
import re
import pandas as pd
import os
from io import BytesIO

############################################
# 1) Utility: extract_text_between
############################################
def extract_text_between(full_text, start_keyword, end_keyword):
    start_index = full_text.find(start_keyword)
    if start_index == -1:
        return ""
    end_index = full_text.find(end_keyword, start_index)
    if end_index == -1:
        return ""
    return full_text[start_index + len(start_keyword):end_index].strip()

############################################
# 2) Utility: format_currency
############################################
def format_currency(val):
    try:
        if isinstance(val, str) and val.startswith("$"):
            num = float(val.replace("$", "").replace(",", ""))
            return f"${num:,.0f}"
        return val
    except:
        return val

############################################
# 3) Utility: unify_multiline_coverage
############################################
def unify_multiline_coverage(df):
    if df.empty or "Coverages" not in df.columns:
        return df
    df = df.fillna("")
    merged_rows = []
    skip_next = False
    for i in range(len(df)):
        if skip_next:
            skip_next = False
            continue
        row = df.iloc[i].copy()
        coverage = row["Coverages"].strip()
        if i < len(df) - 1:
            next_row = df.iloc[i+1]
            next_cov = next_row["Coverages"].strip()
            # If the next row is purely coverage text (no numeric data),
            # merge it into this row’s coverage
            if next_cov and not next_row["Deductible"].strip() \
               and not next_row["Limit"].strip() \
               and not next_row["Premium"].strip():
                coverage = coverage + " " + next_cov
                skip_next = True
        row["Coverages"] = coverage
        merged_rows.append(row)
    return pd.DataFrame(merged_rows, columns=df.columns)

############################################
# 4) Utility: fix_alignment
############################################
def fix_alignment(df):
    df = df.copy()
    for idx, row in df.iterrows():
        ded = row["Deductible"].strip()
        lim = row["Limit"].strip()
        if not lim and ded and re.search(r'\$\d', ded):
            df.at[idx, "Limit"] = ded
            df.at[idx, "Deductible"] = ""
    return df

############################################
# 5) Utility: trim_table_by_keyword
############################################
def trim_table_by_keyword(df, keyword, column="Coverages", include_row=True):
    df = df.copy()
    df["norm"] = df[column].str.strip().str.lower()
    keyword_norm = keyword.strip().lower()
    indices = df.index[df["norm"].str.contains(keyword_norm)]
    if not indices.empty:
        first_idx = indices[0]
        if include_row:
            df = df.loc[:first_idx]
        else:
            df = df.loc[:first_idx-1]
    df.drop(columns=["norm"], inplace=True)
    return df

############################################
# 6) parse_other_coverages_text
############################################
def parse_other_coverages_text(text):
    rows = []
    lines = text.splitlines()
    if lines and "location no./building no" in lines[0].lower():
        lines = lines[1:]
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.lower().startswith("when"):
            continue
        tokens = re.split(r"\s{2,}", line)
        if len(tokens) < 4:
            tokens = re.split(r"\s+", line)

        # Merge tokens that are likely split parts of a numeric value
        merged_tokens = []
        i = 0
        while i < len(tokens):
            token = tokens[i]
            if i + 1 < len(tokens) and re.match(r'^,?\d{3}$', tokens[i+1]):
                token = token + tokens[i+1]
                merged_tokens.append(token)
                i += 2
            else:
                merged_tokens.append(token)
                i += 1
        tokens = merged_tokens

        if len(tokens) < 4:
            continue

        location = tokens[0]
        premium = tokens[-1]
        limit = tokens[-2]
        coverage = " ".join(tokens[1:-2])

        rows.append({
            "Location No./Building No.": location,
            "Coverage": coverage,
            "Limit": limit,
            "Premium": premium
        })

    # If Limit is "$" (or empty) but Coverage ends with a numeric chunk, move it over.
    for row in rows:
        coverage_val = row["Coverage"]
        limit_val = row["Limit"]
        match = re.search(r'(.*)\s+(\$?\d[\d,]+)$', coverage_val)
        if match and (limit_val == "" or limit_val == "$"):
            row["Coverage"] = match.group(1)
            row["Limit"] = match.group(2)

    return pd.DataFrame(rows, columns=["Location No./Building No.", "Coverage", "Limit", "Premium"])

############################################
# 7) parse_other_coverages_pdfplumber
############################################
def parse_other_coverages_pdfplumber(pdf_path):
    page_num = None
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        for i, page in enumerate(pdf.pages, start=1):
            txt = page.extract_text() or ""
            if re.search(r"OTHER\s*COVERAGES", txt, re.IGNORECASE):
                page_num = i
                break
    if page_num is None:
        return pd.DataFrame()

    lines_to_parse = []
    capturing = False
    with pdfplumber.open(pdf_path) as pdf:
        for p in range(page_num - 1, total_pages):
            page_text = pdf.pages[p].extract_text() or ""
            lines = page_text.split("\n")
            for line in lines:
                low_line = line.lower()
                if re.search(r"other\s*coverages", low_line):
                    capturing = True
                    continue
                if "mortgage holder(s)" in low_line:
                    capturing = False
                if capturing:
                    lines_to_parse.append(line)
            if not capturing and lines_to_parse:
                break

    text_to_parse = "\n".join(lines_to_parse)
    df = parse_other_coverages_text(text_to_parse)
    return df

############################################
# 8) parse_property_coverages
############################################
def parse_property_coverages(section_text, full_text, proposal_index):
    rows = []
    lines = section_text.splitlines()
    for line in lines:
        l = line.strip()
        if not l:
            continue
        if l.upper() == "PREMIUM":
            continue
        if "total quote premium" in l.lower():
            m = re.search(r"\$\s*\d[\d,\.]*", l, re.IGNORECASE)
            if m:
                rows.append({"Coverage": "Total Quote Premium", "Premium": m.group(0).replace(" ", "")})
            else:
                rows.append({"Coverage": "Total Quote Premium", "Premium": ""})
            break
        parts = l.rsplit(" ", 1)
        if len(parts) == 2:
            coverage = parts[0].strip()
            premium = parts[1].strip()
            if coverage.endswith("$"):
                coverage = coverage[:-1].strip()
                if premium and premium != "•" and not premium.startswith("$"):
                    premium = "$" + premium
            else:
                if premium and premium != "•" and not premium.startswith("$"):
                    premium = "$" + premium
            rows.append({"Coverage": coverage, "Premium": premium})
        else:
            rows.append({"Coverage": l, "Premium": ""})
    if not any(r["Coverage"].lower() == "total quote premium" for r in rows):
        rows.append({"Coverage": "Total Quote Premium", "Premium": ""})
    return rows

############################################
# 9) parse_premises_into_blocks (not used)
############################################
def parse_premises_into_blocks(premises_text):
    return []

############################################
# 10) display_coverage_blocks_as_tables (not used)
############################################
def display_coverage_blocks_as_tables(blocks):
    return

############################################
#  HELPER: extract_text_pdfplumber
############################################
def extract_text_pdfplumber(pdf_bytes: bytes, use_layout: bool = True) -> str:
    try:
        import pdfplumber
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            if use_layout:
                text = "\n".join(
                    page.extract_text(layout=True) or "" for page in pdf.pages
                )
            else:
                text = "\n".join(
                    page.extract_text() or "" for page in pdf.pages
                )
        return text
    except Exception as e:
        print(f"Error extracting text: {e}")
        return ""

def clean_string_for_compare(s: str) -> str:
    return re.sub(r'[^A-Za-z0-9]+', '', s).upper()

def is_limit_or_premium_candidate(token: str) -> bool:
    if token.lower() == "included":
        return True
    if token.startswith("$"):
        return True
    try:
        float(token.replace(",", ""))
        return True
    except ValueError:
        return False

############################################
#  parse_policy_endorsements_table
############################################

def parse_policy_endorsements_table(text: str) -> pd.DataFrame:
    import re
    num_pattern = re.compile(r'\$\d[\d,\.]*%?|\d[\d,\,]*\s*(?:Feet|Days)|\d+%?|Included', re.IGNORECASE)

    # 1) Split into non-empty lines
    lines = [ln for ln in text.splitlines() if ln.strip()]
    start = None
    for i, ln in enumerate(lines):
        c = clean_string_for_compare(ln)
        if c.startswith("POLICYLEVEL") and ("ENDORSEMENT" in c or "COVERAGE" in c):
            start = i
            break
    if start is None:
        return pd.DataFrame(columns=["Coverages","Deductible","Limit","Premium"])

    # 2) Collect lines after header row
    raw = []
    collecting = False
    for ln in lines[start+1:]:
        clean_ln = clean_string_for_compare(ln)
        if clean_ln.startswith("OTHERCOVERAGES"):
            break
        if not collecting:
            if all(h in clean_ln for h in ("COVERAGES","DEDUCTIBLE","LIMIT","PREMIUM")):
                collecting = True
            continue
        if ln.strip().upper().startswith("PROPOSAL"):
            continue
        raw.append(ln.strip())

    # 3) Merge wrapped lines: hyphen wrap, unmatched parentheses, or 'or ' continuation
    merged = []
    i = 0
    while i < len(raw):
        cur = raw[i]
        nxt = raw[i+1] if i+1 < len(raw) else ""
        nums_cur = num_pattern.findall(cur)
        nums_nxt = num_pattern.findall(nxt)
        merge_cond = False
        if cur.endswith('-'):
            merge_cond = True
        elif cur.count('(') > cur.count(')'):
            merge_cond = True
        elif nxt.lower().startswith('or '):
            merge_cond = True
        if merge_cond:
            merged.append(cur + ' ' + nxt)
            i += 2
        else:
            merged.append(cur)
            i += 1

    # 4) Parse each merged line
    rows = []
    for raw_line in merged:
        nums = num_pattern.findall(raw_line)
        m = num_pattern.search(raw_line)
        cov = raw_line[:m.start()].strip() if m else raw_line.strip()

        # Handle 'Endorsement' parent+child
        if "Endorsement" in cov and nums:
            parts = cov.split("Endorsement", 1)
            parent = parts[0].strip() + " Endorsement"
            child  = parts[1].strip()
            limit = nums[0]
            prem = "Included"
            rows.append({"Coverages": parent, "Deductible": "", "Limit": "", "Premium": ""})
            rows.append({"Coverages": child,  "Deductible": "", "Limit": limit, "Premium": prem})
            continue

        # Standard rows: first numeric as Limit, Premium as 'Included'
        limit = nums[0] if nums else ""
        prem = "Included" if nums else ""
        # Exception for first row
        if not rows and nums:
            rows.append({"Coverages": cov, "Deductible": "", "Limit": "", "Premium": limit})
        else:
            rows.append({"Coverages": cov, "Deductible": "", "Limit": limit, "Premium": prem})

    return pd.DataFrame(rows, columns=["Coverages","Deductible","Limit","Premium"])


############################################
#  parse_policy_forms
############################################


def parse_policy_endorsements_table_old(text: str) -> pd.DataFrame:
    all_lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    start_idx = None
    for i, line in enumerate(all_lines):
        if any(k in clean_string_for_compare(line) for k in ("POLICYLEVELENDORSEMENTS","POLICYLEVELCOVERAGES")):
            start_idx = i
            break
    if start_idx is None:
        return pd.DataFrame(columns=["Coverages", "Deductible", "Limit", "Premium"])

    required_headers_cleaned = {"COVERAGES", "DEDUCTIBLE", "LIMIT", "PREMIUM"}
    found_headers = set()
    data_lines = []
    collecting = False

    for line in all_lines[start_idx + 1:]:
        if "OTHERCOVERAGES" in clean_string_for_compare(line):
            break
        if not collecting:
            line_clean = clean_string_for_compare(line)
            if "COVERAGES" in line_clean:
                found_headers.add("COVERAGES")
            if "DEDUCTIBLE" in line_clean:
                found_headers.add("DEDUCTIBLE")
            if "LIMIT" in line_clean:
                found_headers.add("LIMIT")
            if "PREMIUM" in line_clean:
                found_headers.add("PREMIUM")
            if found_headers == required_headers_cleaned:
                collecting = True
            continue
        else:
            data_lines.append(line)

    combined_text = " ".join(data_lines)
    tokens = combined_text.split()

    rows = []
    i = 0
    while i < len(tokens):
        coverage_tokens = []
        while i < len(tokens) and not is_limit_or_premium_candidate(tokens[i]):
            coverage_tokens.append(tokens[i])
            i += 1
        coverage_str = " ".join(coverage_tokens)

        numeric_candidates = []
        while i < len(tokens) and is_limit_or_premium_candidate(tokens[i]):
            numeric_candidates.append(tokens[i])
            i += 1
            if numeric_candidates[-1] == "$" and i < len(tokens) and is_limit_or_premium_candidate(tokens[i]):
                numeric_candidates[-1] = "$" + tokens[i].lstrip("$")
                i += 1

        ded_str, lim_str, prem_str = "", "", ""
        if len(numeric_candidates) == 3:
            ded_str, lim_str, prem_str = numeric_candidates
        elif len(numeric_candidates) == 2:
            lim_str, prem_str = numeric_candidates
        elif len(numeric_candidates) == 1:
            prem_str = numeric_candidates[0]

        rows.append({
            "Coverages": coverage_str,
            "Deductible": ded_str,
            "Limit": lim_str,
            "Premium": prem_str
        })

    return pd.DataFrame(rows, columns=["Coverages", "Deductible", "Limit", "Premium"])

############################################
#  parse_policy_forms
############################################



def parse_policy_endorsements_table_combined(text: str) -> pd.DataFrame:
    """
    Selects parser based on PDF version detected by header.
    If the header contains 'POLICY LEVEL COVERAGES', uses the new parser; otherwise uses the old parser.
    """
    upper_text = text.upper()
    if "POLICY LEVEL COVERAGES" in upper_text:
        df = parse_policy_endorsements_table(text)
    else:
        df = parse_policy_endorsements_table_old(text)
    df = fix_alignment(df)
    return df


def parse_policy_forms(text: str) -> dict:
    coverage_titles = [
        "Commercial Property Coverage Part",
        "Commercial Property Forms"
    ]
    start_index = text.find("SCHEDULE OF FORMS AND ENDORSEMENTS")
    if start_index == -1:
        return {}
    relevant_lines = text[start_index:].splitlines()
    sections = {}
    current_title = None
    i = 0
    while i < len(relevant_lines):
        line = relevant_lines[i].strip()
        if line in coverage_titles:
            current_title = line
            sections[current_title] = []
            i += 1
            if i < len(relevant_lines):
                possible_header_line = relevant_lines[i].lower().strip()
                if ("number" in possible_header_line and
                    "edition" in possible_header_line and
                    "description" in possible_header_line):
                    i += 1
            continue
        if current_title:
            if line.startswith("Commercial ") and (line not in coverage_titles):
                current_title = None
                i -= 1
            else:
                row = parse_line_into_columns(line)
                number, edition, description = row
                if not edition:
                    if sections[current_title]:
                        sections[current_title][-1][2] += " " + " ".join(
                            part for part in [number, description] if part
                        )
                        sections[current_title][-1][2] = sections[current_title][-1][2].strip()
                    else:
                        sections[current_title].append(["", "", number + " " + description])
                else:
                    sections[current_title].append(row)
        i += 1
    return sections

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

############################################
# 14) HELPER: Make a DataFrame's HTML cells editable
############################################
def make_table_cells_editable(html_str: str) -> str:
    pattern = r'(<td)([^>]*>)'
    replace = r'<td contenteditable="true"\2'
    return re.sub(pattern, replace, html_str, flags=re.IGNORECASE)

##############################################################
#  parse_property_pdf(pdf_bytes) - unified extraction
##############################################################
def parse_property_pdf(pdf_bytes):
    """
    Reads the PDF in memory, extracts all Property data,
    and returns the DataFrames as a dictionary.
    This is the same logic that was in your standalone main().
    """
    import pdfplumber

    temp_pdf_path = "temp_forced_4cols.pdf"
    with open(temp_pdf_path, "wb") as f:
        f.write(pdf_bytes)

    with pdfplumber.open(temp_pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            txt = page.extract_text() or ""
            full_text += txt + "\n"

    # 1) PROPERTY COVERAGES
    proposal_index = 0
    coverages_section = extract_text_between(
        full_text[proposal_index:],
        "PROPERTY COVERAGES",
        "DESCRIPTION OF PREMISES AND COVERAGES PROVIDED"
    )
    df_cov = pd.DataFrame()
    if coverages_section:
        coverage_rows = parse_property_coverages(coverages_section, full_text, proposal_index)
        df_cov = pd.DataFrame(coverage_rows)

    # 2) DESCRIPTION OF PREMISES
    section_start = full_text.find("DESCRIPTION OF PREMISES AND COVERAGES PROVIDED")
    section_text = ""
    if section_start != -1:
        section_text = full_text[section_start:]
        # Trim at first of multiple end markers
        end_markers = [
            "Equipment Breakdown",
            "POLICY LEVEL ENDORSEMENTS",
            "Policy Level Coverages", "POLICY LEVEL COVERAGES",
            "Policy Level Endorsements",
            "OTHER COVERAGES"
        ]
        end_positions = [section_text.find(m) for m in end_markers if section_text.find(m) != -1]
        if end_positions:
            cut_at = min(end_positions)
            section_text = section_text[:cut_at]
    # Blanket Coverages
    df_blanket = pd.DataFrame()
    def parse_blanket_coverages(text: str):
        blanket_rows = []
        blanket_dict = {}
        pattern = re.compile(r"(?i)^(Blanket\s+[^\n]*?)([\d,]{3,})", re.MULTILINE)
        for match in pattern.finditer(text):
            full_type = match.group(1).strip()
            limit_val = match.group(2).strip()
            blanket_rows.append({"Type": full_type, "Limit": limit_val})
        return blanket_rows

    if section_text:
        blanket_rows = parse_blanket_coverages(section_text)
        if blanket_rows:
            df_blanket = pd.DataFrame(blanket_rows, columns=["Type", "Limit"])

    # 3) LOCATION COVERAGES
    df_main = pd.DataFrame(columns=[
        "Loc/Bld", "Address", "Coverage", "Limit",
        "Valuation", "Co-Ins", "Ded", "W/H Ded", "Premium"
    ])
    if section_text:
        blocks = section_text.split("Location No./Building No.")
        if len(blocks) >= 2:
            def combine_multiline_descriptors(lines_list):
                combined = []
                skip_next = False
                for i in range(len(lines_list)):
                    if skip_next:
                        skip_next = False
                        continue
                    line = lines_list[i]
                    if line.strip().endswith('-') and i < len(lines_list) - 1:
                        next_line = lines_list[i+1]
                        merged = line.rstrip('-').strip() + " " + next_line
                        combined.append(merged)
                        skip_next = True
                    else:
                        combined.append(line)
                return combined

            def parse_coverage_line(line: str) -> dict:
                lower_line = line.lower()
                if "wind/hail-ded" in lower_line:
                    return {}
                if "inflation guard" in lower_line:
                    premium_match = re.search(r"(\$?\d[\d,\.]*)", line)
                    premium_val = premium_match.group(1) if premium_match else ""
                    if premium_val and not premium_val.startswith("$"):
                        premium_val = "$" + premium_val
                    return {
                        "Coverage": "Inflation Guard",
                        "Limit": "",
                        "Valuation": "",
                        "Co-Ins": "",
                        "Premium": premium_val
                    }
                if ("ordinary payroll" in lower_line and
                    ("exclusion" in lower_line or "limitation" in lower_line or "applies" in lower_line)):
                    return {}

                match = re.match(r"^(.*?)(\d[\d,]*)(.*)$", line)
                if not match:
                    return {}
                coverage_desc = match.group(1).strip()
                all_nums = re.findall(r"\d[\d,]*", line)
                if not all_nums:
                    return {}
                limit_val = all_nums[0]
                premium_val = all_nums[-1] if len(all_nums) > 1 else ""
                valuation_val = "RC" if "rc" in line.lower() else ""
                pct_match = re.search(r"(\d+%)", line)
                co_ins_val = pct_match.group(1) if pct_match else ""

                return {
                    "Coverage": coverage_desc,
                    "Limit": limit_val,
                    "Valuation": valuation_val,
                    "Co-Ins": co_ins_val,
                    "Premium": premium_val
                }

            all_rows = []
            for idx, raw_block in enumerate(blocks[1:], start=1):
                lines = [ln.strip() for ln in raw_block.splitlines() if ln.strip()]
                if not lines:
                    continue
                first_line = lines[0]
                loc_bld = ""
                ded = ""
                loc_match = re.search(r"(\d{3}/\d{3})", first_line)
                if loc_match:
                    loc_bld = loc_match.group(1)
                else:
                    loc_bld = first_line.strip()
                ded_match = re.search(r"Deductible:\s*(\$[\d,]+)", first_line)
                if ded_match:
                    ded = ded_match.group(1).strip()
                lines.pop(0)
                address_lines = []
                remain_lines = []
                for ln in lines:
                    low_ln = ln.lower()
                    if low_ln.startswith("proposal"):
                        continue  # skip page footer markers that disrupt address mapping
                    if low_ln.startswith("street address"):
                        text = re.sub(r"(?i)^street address\s*", "", ln)
                        text = re.sub(r"(?i)(\d+\s*story.*|joisted\s*masonry.*|occupied\s*as.*|wind/hail-ded.*)", "", text)
                        address_lines.append(text.strip())
                    elif low_ln.startswith("city, state and zip code"):
                        pass
                    else:
                        remain_lines.append(ln)
                address = " ".join(filter(None, address_lines))
                lines = combine_multiline_descriptors(remain_lines)

                block_text = "\n".join(lines)
                wh_ded_match = re.search(r"Wind/Hail[- ]Ded:\s*(\d+%?)", block_text, re.IGNORECASE)
                wh_ded = wh_ded_match.group(1).strip() if wh_ded_match else ""

                coverage_lines = []
                i = 0
                while i < len(lines):
                    current_line = lines[i]
                    if "business income" in current_line.lower() and not re.search(r"(\d[\d,]*)", current_line):
                        combined_line = current_line
                        i += 1
                        while i < len(lines) and not re.search(r"(\d[\d,]*)", lines[i]):
                            combined_line += " " + lines[i]
                            i += 1
                        if i < len(lines):
                            combined_line += " " + lines[i]
                            i += 1
                        parsed = parse_coverage_line(combined_line)
                        # --- business-income override: split out Actual Loss Sustained and premium into correct columns
                        if parsed and parsed.get('Coverage', '').lower().startswith('business income'):
                            raw_line = combined_line
                            # extract premium value
                            m_prem = re.search(r'\$\s*([\d,]+)', raw_line)
                            premium_raw = f"${m_prem.group(1)}" if m_prem else parsed.get('Premium', '')
                            # remove premium and SPECIAL from coverage part
                            cov_part = raw_line
                            if m_prem:
                                cov_part = cov_part.replace(m_prem.group(0), '')
                            cov_part = re.sub(r'\bSPECIAL\b', '', cov_part, flags=re.IGNORECASE).strip()
                            # split out Actual Loss Sustained if present
                            if re.search(r'Actual Loss Sustained', cov_part, re.IGNORECASE):
                                parts = re.split(r'(Actual Loss Sustained)', cov_part, flags=re.IGNORECASE)
                                coverage_text = parts[0].strip()
                                limit_text = parts[1].strip()
                            else:
                                coverage_text = cov_part
                                limit_text = ''
                            # override parsed fields
                            parsed['Coverage'] = coverage_text
                            parsed['Limit'] = limit_text
                            parsed['Premium'] = premium_raw
                        # --- end business-income override ---
                    else:
                        parsed = parse_coverage_line(current_line)
                        i += 1

                    if parsed:
                        coverage_name_lower = parsed["Coverage"].lower()
                        if coverage_name_lower == "business income":
                            used_ded = "Waiting Period"
                            used_wh_ded = ""
                        elif coverage_name_lower == "inflation guard":
                            used_ded = ""
                            used_wh_ded = ""
                        else:
                            used_ded = ded
                            used_wh_ded = wh_ded

                        if "blanket" in coverage_name_lower:
                            parsed["Coverage"] = re.sub(r"(?i)\s*blanket.*", "", parsed["Coverage"]).strip()
                            parsed["Limit"] = "See Blanket"

                        row = {
                            "Loc/Bld": loc_bld,
                            "Address": "",
                            "Coverage": parsed["Coverage"],
                            "Limit": parsed["Limit"],
                            "Valuation": parsed["Valuation"],
                            "Co-Ins": parsed["Co-Ins"],
                            "Ded": used_ded,
                            "W/H Ded": used_wh_ded,
                            "Premium": parsed["Premium"]
                        }
                        coverage_lines.append(row)

                if coverage_lines:
                    coverage_lines[0]["Address"] = address
                    for j in range(1, len(coverage_lines)):
                        coverage_lines[j]["Loc/Bld"] = ""
                        coverage_lines[j]["Address"] = ""
                    all_rows.extend(coverage_lines)

            filtered_rows = []
            for row in all_rows:
                coverage_name_lower = row["Coverage"].strip().lower()
                loc_lower = row["Loc/Bld"].strip().lower()
                if "equipment breakdown" in coverage_name_lower and loc_lower != "all locations":
                    continue
                filtered_rows.append(row)
            all_rows = filtered_rows

            for r in all_rows:
                if r["Limit"] and r["Limit"] != "See Blanket" and not r["Limit"].startswith("$"):
                    r["Limit"] = "$" + r["Limit"]
                if r["Premium"] and not r["Premium"].startswith("$"):
                    r["Premium"] = "$" + r["Premium"]

            df_main = pd.DataFrame(all_rows, columns=[
                "Loc/Bld", "Address", "Coverage", "Limit",
                "Valuation", "Co-Ins", "Ded", "W/H Ded", "Premium"
            ])

    # --- post-process location coverages: remove PROPOSAL rows and clear deductibles for business income
    if not df_main.empty:
        # drop any stray 'PROPOSAL' entries
        df_main = df_main[df_main['Coverage'].str.strip().str.upper() != 'PROPOSAL']
        # clear Ded and W/H Ded for Business Income rows
        mask_bi = df_main['Coverage'].str.lower().str.startswith('business income')
        df_main.loc[mask_bi, ['Ded', 'W/H Ded']] = ''
        



        # remove leading $ from Actual Loss Sustained in Limit
        mask_act = df_main['Limit'].str.lower().str.startswith('$actual loss')
        df_main.loc[mask_act, 'Limit'] = df_main.loc[mask_act, 'Limit'].str.lstrip('$')
    # 4) POLICY LEVEL ENDORSEMENTS
    df_endorsements = pd.DataFrame()
    endorsements_page = None
    with pdfplumber.open(temp_pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            txt = page.extract_text() or ""
            if "POLICY LEVEL ENDORSEMENTS" in txt.upper() or "POLICY LEVEL COVERAGES" in txt.upper():
                endorsements_page = i
                break
    if endorsements_page is not None:
        with pdfplumber.open(temp_pdf_path) as pdf:
            endorsements_text = ""
            for i in range(endorsements_page - 1, min(endorsements_page - 1 + 2, len(pdf.pages))):
                endorsements_text += pdf.pages[i].extract_text() + "\n"
        tmp_end = parse_policy_endorsements_table_combined(endorsements_text)
        if not tmp_end.empty:
            for col in ["Deductible", "Limit", "Premium"]:
                tmp_end[col] = tmp_end[col].apply(format_currency)
            df_endorsements = tmp_end[["Coverages", "Deductible", "Limit", "Premium"]]

    # 5) OTHER COVERAGES
    df_other = pd.DataFrame()
    with pdfplumber.open(temp_pdf_path) as pdf:
        # same approach as parse_other_coverages_pdfplumber
        df_tmp = parse_other_coverages_pdfplumber(temp_pdf_path)
        if not df_tmp.empty:
            for col in ["Limit", "Premium"]:
                df_tmp[col] = df_tmp[col].apply(format_currency)
            df_other = df_tmp

    # 6) POLICY FORMS
    forms_sections = parse_policy_forms(full_text)

    # Clean up
    if os.path.exists(temp_pdf_path):
        os.remove(temp_pdf_path)

    # Return all data as a dictionary
    return {
        "df_cov": df_cov,                # Property Coverages
        "df_blanket": df_blanket,        # Blanket Coverages
        "df_main": df_main,              # Location Coverages
        "df_endorsements": df_endorsements, 
        "df_other": df_other,            # Other Coverages
        "forms_sections": forms_sections # Dict of forms
    }

# If you still want to run Property.py standalone for debugging:
if __name__ == "__main__":
    print("Run this file from Main.py or import parse_property_pdf in your code.")
