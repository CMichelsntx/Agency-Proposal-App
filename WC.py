import streamlit as st
import io
import re
import pandas as pd

#########################################
# 1) Global CSS to center table headers
#########################################
st.markdown("""
<style>
thead tr th {
    text-align: center !important;
}
</style>
""", unsafe_allow_html=True)

#########################################
# 2) Make table cells editable
#########################################
def make_table_cells_editable(html_str: str) -> str:
    """
    Inserts contenteditable="true" into each <td> so that the user can edit the cell content.
    """
    pattern = r'(<td)([^>]*>)'
    replace = r'<td contenteditable="true"\2'
    return re.sub(pattern, replace, html_str, flags=re.IGNORECASE)

#########################################
# 3) Fix split notice lines
#########################################
def fix_split_notice_lines(lines):
    """
    Merges lines if we detect a split notice like:
       'Motor Vehicle Crime Prevention Authority Fee (See enclosed'
       + 'notice):'
    """
    merged = []
    skip_next = False
    for i in range(len(lines)):
        if skip_next:
            skip_next = False
            continue
        if i < len(lines) - 1 and lines[i+1].lower() == "notice):":
            combined = lines[i] + " notice):"
            merged.append(combined)
            skip_next = True
        else:
            merged.append(lines[i])
    return merged

#########################################
# 4) PDFMiner-based text extraction
#########################################
def get_pdf_lines(file_bytes):
    """
    Use pdfminer.six to extract text from the PDF and return as a list of lines.
    """
    from pdfminer.high_level import extract_text
    text = extract_text(io.BytesIO(file_bytes))
    # Split on newlines and strip
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    # Fix certain split lines
    lines = fix_split_notice_lines(lines)
    return lines

#########################################
# Additional PDF extraction debug functions
#########################################
def get_pdf_text_pypdf2(file_bytes):
    """
    Use PyPDF2 to extract text from the PDF.
    """
    import PyPDF2
    reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
    text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text

def get_pdf_text_pdfplumber(file_bytes):
    """
    Use pdfplumber to extract text from the PDF.
    """
    import pdfplumber
    text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text

def get_pdf_text_pymupdf(file_bytes):
    """
    Use PyMuPDF (fitz) to extract text from the PDF.
    """
    import fitz  # PyMuPDF
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text() + "\n"
    return text

#########################################
# Camelot Stream Debug Function
#########################################
def get_camelot_stream_tables(file_bytes):
    """
    Uses Camelot in stream mode to extract tables.
    Writes file_bytes to a temporary file and returns the Camelot tables.
    """
    import camelot
    import tempfile
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(file_bytes)
        tmp.flush()
        tables = camelot.read_pdf(tmp.name, flavor="stream")
    return tables

#########################################
# 5) Policy Information Extraction (modified)
#########################################
def extract_policy_information(lines):
    def clean_colon_space(s):
        return re.sub(r'^[:\s]+', '', s).strip()

    result = {
        "Date": "",
        "Rating Company": "",  # New key added
        "Quote No.": "",
        "Policy No.": "",
        "NCCI Carrier Code No.": "",
        "FEIN": "",
        "Risk ID No.": "",
        "Bureau File No.": "",
        "Entity of Insured": "",
        "Proposed Policy Period": "",
        "Named Insured": "",
        "DBA": "",
        "Insured Address": "",
        "Insured City, State & Zip": "",
        "Agent Name": "",
        "Agent Phone": "",
        "Agent Address": "",
        "Agent City, State & Zip": ""
    }

    combined_text = "\n".join(lines)

    # 1) Date
    date_match = re.search(r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b', combined_text)
    if date_match:
        result["Date"] = date_match.group(0)

    # New: Rating Company extraction
    rating_match = re.search(r'Rating Company:\s*(.*)', combined_text, re.IGNORECASE)
    if rating_match:
        result["Rating Company"] = rating_match.group(1).strip()

    # 2) Proposed Policy Period
    period_match = re.search(r"The Proposed Policy Period is from\s*(.*?)\s+at", combined_text, re.IGNORECASE)
    if period_match:
        result["Proposed Policy Period"] = period_match.group(1).strip()

    # 3) Named Insured, DBA, Address
    for i, line in enumerate(lines):
        if "Named Insured Name and Address" in line:
            if i + 1 < len(lines):
                result["Named Insured"] = lines[i + 1]
            if i + 2 < len(lines):
                if "dba" in lines[i + 2].lower():
                    result["DBA"] = lines[i + 2]
                    if i + 3 < len(lines):
                        result["Insured Address"] = lines[i + 3]
                    if i + 4 < len(lines):
                        result["Insured City, State & Zip"] = lines[i + 4]
                else:
                    result["Insured Address"] = lines[i + 2]
                    if i + 3 < len(lines):
                        result["Insured City, State & Zip"] = lines[i + 3]
            break

    # 4) Agent Info
    for i, line in enumerate(lines):
        if "Agency Name and Address" in line:
            if i + 1 < len(lines):
                result["Agent Phone"] = lines[i + 1]
            if i + 2 < len(lines):
                result["Agent Name"] = lines[i + 2]
            if i + 3 < len(lines):
                candidate = lines[i + 3]
                if candidate and (candidate[0].isdigit() or candidate.upper().startswith("PO")):
                    result["Agent Address"] = candidate
                    if i + 4 < len(lines):
                        result["Agent City, State & Zip"] = lines[i + 4]
            break

    # 5) Policy No. (around "PREMIUM SUMMARY")
    premium_idx = None
    for i, line in enumerate(lines):
        if "PREMIUM SUMMARY" in line.upper():
            premium_idx = i
            break
    if premium_idx is not None:
        for line in lines[premium_idx:]:
            if re.search(r"Policy No\.?", line, re.IGNORECASE):
                parts = re.split(r'Policy No\.?\s*', line, flags=re.IGNORECASE)
                if len(parts) > 1:
                    result["Policy No."] = clean_colon_space(parts[1])
                break

    if not result["Policy No."]:
        for line in lines:
            match = re.search(r'Policy No\.?:?\s*(.*)', line, re.IGNORECASE)
            if match:
                val = match.group(1).strip()
                if val:
                    result["Policy No."] = val
                    break

    # 6) INFORMATION PAGE block (Quote No., NCCI Carrier)
    info_page_idx = None
    for i, line in enumerate(lines):
        if "INFORMATION PAGE" in line.upper():
            info_page_idx = i
            break

    if info_page_idx is not None:
        quote_pattern = re.compile(r'Quote No\.?:?\s*(.*)', re.IGNORECASE)
        ncci_pattern = re.compile(r'NCCI Carrier Code No\.?:?\s*(.*)', re.IGNORECASE)
        for line in lines[info_page_idx:]:
            q_match = quote_pattern.search(line)
            if q_match:
                result["Quote No."] = q_match.group(1).strip()
            ncci_match = ncci_pattern.search(line)
            if ncci_match:
                result["NCCI Carrier Code No."] = ncci_match.group(1).strip()

    if not result["Quote No."]:
        for line in lines:
            match = re.search(r'Quote No\.?:?\s*(.*)', line, re.IGNORECASE)
            if match:
                result["Quote No."] = match.group(1).strip()
                break

    # 7) "Refer to Name and Location Schedule" block
    refer_idx = None
    for i, line in enumerate(lines):
        if "refer to name and location schedule" in line.lower():
            refer_idx = i
            break

    if refer_idx is not None:
        i = refer_idx
        while i < len(lines):
            line = lines[i]
            fein_match = re.search(r'(?i)FEIN:\s*(.*)', line)
            if fein_match:
                raw_fein = fein_match.group(1).strip()
                numeric_match = re.search(r'(\d+)', raw_fein)
                if numeric_match:
                    result["FEIN"] = numeric_match.group(1)
                else:
                    result["FEIN"] = raw_fein
            riskid_match = re.search(r'(?i)Risk ID No\.?:?\s*(.*)', line)
            if riskid_match:
                result["Risk ID No."] = riskid_match.group(1).strip()
            bureau_match = re.search(r'(?i)Bureau File No\.?:?\s*(.*)', line)
            if bureau_match:
                result["Bureau File No."] = bureau_match.group(1).strip()
            entity_match = re.search(r'(?i)Entity of Insured:?\s*(.*)', line)
            if entity_match:
                entity_data = entity_match.group(1).strip()
                if (i + 1) < len(lines):
                    next_line = lines[i+1]
                    if (":" not in next_line) and ("No." not in next_line) and next_line.strip():
                        entity_data += " " + next_line.strip()
                result["Entity of Insured"] = entity_data
            i += 1

    return result

#########################################
# 6) Extract Coverages (unchanged)
#########################################
def extract_coverages(lines):
    coverage_start = None
    for i, line in enumerate(lines):
        if "COVERAGE INFORMATION" in line.upper():
            coverage_start = i
            break
    if coverage_start is None:
        return [], []

    coverage_block = []
    premium_index = None
    for j in range(coverage_start + 1, len(lines)):
        if lines[j].strip().upper() == "PREMIUM":
            premium_index = j
            break
        coverage_block.append(lines[j])

    if coverage_block and coverage_block[0].strip().upper() == "COVERAGES":
        coverage_block = coverage_block[1:]

    premiums = []
    if premium_index is not None:
        idx = premium_index + 1
        while idx < len(lines) and lines[idx].strip() == "$":
            idx += 1
        while idx < len(lines):
            if re.search(r'\d', lines[idx]):
                premiums.append(lines[idx])
            idx += 1

    if len(premiums) < len(coverage_block):
        premiums += [""] * (len(coverage_block) - len(premiums))
    elif len(premiums) > len(coverage_block):
        premiums = premiums[:len(coverage_block)]

    return coverage_block, premiums

#########################################
# 7) Extract Workers Comp Table (unchanged)
#########################################
def extract_workers_comp_table(lines):
    start_idx = None
    for i, line in enumerate(lines):
        if "WORKERS COMPENSATION AND EMPLOYERS LIABILITY QUOTE PROPOSAL" in line.upper():
            start_idx = i
            break
    if start_idx is None:
        return []

    coverage_idx = None
    for i in range(start_idx, len(lines)):
        if "COVERAGE" in lines[i].upper():
            coverage_idx = i
            break
    if coverage_idx is None:
        return []
    
    extracted_rows = []
    pattern = re.compile(r'(.*?)(\$\s*[\d,]+)(.*)', re.IGNORECASE)
    for i in range(coverage_idx + 1, len(lines)):
        if "PREMIUM" in lines[i].upper():
            break
        match = pattern.search(lines[i])
        if match:
            coverage_text = match.group(1).strip()
            limit_text = match.group(2).strip()
            limit_text = re.sub(r'\s+', '', limit_text)
            type_text = match.group(3).strip()
            extracted_rows.append((coverage_text, limit_text, type_text))
    return extracted_rows

#########################################
# 7a) Table 3 using PDFPlumber (unchanged)
#########################################
def extract_table_3_pdfplumber(file_bytes):
    """
    This is unchanged from your previously perfect version.
    """
    import pdfplumber
    text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    lines = text.splitlines()
    
    premium_idx = None
    est_annual_idx = None
    for i, line in enumerate(lines):
        if premium_idx is None and "PREMIUM" in line.upper():
            premium_idx = i
        if premium_idx is not None and "EST ANNUAL" in line.upper():
            est_annual_idx = i
            break
    if est_annual_idx is None:
        return []
    
    extracted_rows = []
    i = est_annual_idx + 1
    while i < len(lines):
        if "WORKERS COMPENSATION AND EMPLOYERS" in lines[i].upper():
            break
        current_line = lines[i].strip()
        if current_line:
            parts = re.split(r'(\d[\d,]*(?:\.\d+)?)', current_line)
            parts = [p for p in parts if p]  # remove empty strings
            j = 0
            while j < len(parts) - 1:
                desc = parts[j].strip()
                premium_raw = parts[j+1].strip()

                desc = desc.replace("StandardPremium", "Standard Premium").replace("$", "")
                negative_premium = False
                if desc.endswith("-"):
                    negative_premium = True
                    desc = desc[:-1].strip()

                try:
                    val = float(premium_raw.replace(",", ""))
                    if negative_premium:
                        val = abs(val)
                    if val.is_integer():
                        formatted_premium = f"${int(val):,}"
                    else:
                        formatted_premium = f"${val:,.2f}"
                except:
                    formatted_premium = premium_raw

                if negative_premium:
                    formatted_premium = f"-{formatted_premium}"

                if desc and premium_raw:
                    extracted_rows.append((desc, formatted_premium))
                j += 2
        i += 1
    return extracted_rows

#########################################
# 8) Extract Schedule of Operations (unchanged)
#########################################
def extract_schedule_operations_table(lines):
    """
    1) Find "SCHEDULE OF OPERATIONS"
    2) Find the header row ("Loc ST No. Classification...")
    3) Gather lines until "Subtotal" or end.
    4) Each row starts with something like "5 LA 8810" or "5 LA If"
       We accumulate subsequent lines until the next row start or "Subtotal".
    5) Then parse from left to right:
       - loc = tokens[0]
       - st  = tokens[1]
       - code= tokens[2]
       - classification = everything until we see numeric or "If" => that's Premium Basis
       - next => Rate
       - skip possible "$"
       - next => Premium
    """
    # 1) Locate "SCHEDULE OF OPERATIONS"
    schedule_start = None
    for i, line in enumerate(lines):
        if "SCHEDULE OF OPERATIONS" in line.upper():
            schedule_start = i
            break
    if schedule_start is None:
        return [], None

    # 2) Find the header row
    header_idx = None
    for i in range(schedule_start, len(lines)):
        hdr_line = lines[i].upper()
        if "LOC" in hdr_line and "ST" in hdr_line and "NO." in hdr_line and "CLASSIFICATION" in hdr_line:
            header_idx = i
            break
    if header_idx is None:
        return [], None

    # 3) Gather lines until "Subtotal"
    data_lines = []
    subtotal_line = None
    for i in range(header_idx + 1, len(lines)):
        if "SUBTOTAL" in lines[i].upper():
            subtotal_line = lines[i]
            break
        data_lines.append(lines[i])

    # A new row starts if line matches "<digits> <two letters> <digits or 'If'>"
    row_start_pattern = re.compile(r'^(\d+)\s+[A-Za-z]{2}\s+(?:\d+|If)', re.IGNORECASE)

    table_rows = []
    current_row = []

    def finalize_row(row_lines):
        row_str = " ".join(row_lines).strip()
        tokens = row_str.split()
        if len(tokens) < 4:
            return
        
        loc = tokens[0]
        st_val = tokens[1]
        code_val = tokens[2]

        idx = 3
        classification_tokens = []
        while idx < len(tokens):
            if re.match(r'^[\d,]+$', tokens[idx], re.IGNORECASE) or tokens[idx].lower() == "if":
                break
            classification_tokens.append(tokens[idx])
            idx += 1
        classification_str = " ".join(classification_tokens)

        if idx >= len(tokens):
            return
        remuneration_raw = tokens[idx]
        idx += 1
        if remuneration_raw.lower() == "if" and idx < len(tokens) and tokens[idx].lower() == "any":
            remuneration_raw += " " + tokens[idx]
            idx += 1

        if idx >= len(tokens):
            return
        rate_raw = tokens[idx]
        idx += 1

        if idx < len(tokens) and tokens[idx] == "$":
            idx += 1

        if idx >= len(tokens):
            return
        premium_raw = tokens[idx]

        if re.match(r'^[\d,]+$', remuneration_raw):
            remuneration = f"${remuneration_raw}"
        else:
            remuneration = remuneration_raw

        if re.match(r'^[\d.]+$', rate_raw):
            try:
                r_val = float(rate_raw)
                rate = f"${r_val:.2f}"
            except:
                rate = rate_raw
        else:
            rate = rate_raw

        pr_str = premium_raw.replace(",", "")
        try:
            valf = float(pr_str)
            if valf.is_integer():
                premium = f"${int(valf):,}"
            else:
                premium = f"${valf:,.2f}"
        except:
            premium = premium_raw

        table_rows.append((loc, st_val, code_val, classification_str, remuneration, rate, premium))

    for line in data_lines:
        line_clean = line.strip()
        if row_start_pattern.match(line_clean):
            if current_row:
                finalize_row(current_row)
            current_row = [line_clean]
        else:
            if current_row:
                current_row.append(line_clean)

    if current_row:
        finalize_row(current_row)

    sub_data = None
    if subtotal_line:
        m = re.search(r'Subtotal:\s*(.*?)\s+\$\s*([\d,]+)', subtotal_line, re.IGNORECASE)
        if m:
            text_part = m.group(1).strip()
            amt_part = "$" + m.group(2)
            sub_data = ("Subtotal", text_part, amt_part)

    return table_rows, sub_data

#########################################
# NEW: Additional Premium Info Extraction (unchanged)
#########################################
def extract_additional_premium_info(lines):
    """
    A simpler approach to parse additional premium info rows after "Subtotal:".
    1) Capture lines until "WORKERS COMPENSATION AND EMPLOYERS".
    2) Accumulate lines until we see one that has a "$" => we treat that as the end of a row.
    3) Use a single regex to parse Code No., Description, and Premium from each row.
    4) Special-case "Total State Standard Premium".
    """
    import re

    start_index = None
    for i, line in enumerate(lines):
        if "Subtotal:" in line:
            start_index = i
            break
    if start_index is None:
        return []

    section_lines = []
    for line in lines[start_index + 1:]:
        if "WORKERS COMPENSATION AND EMPLOYERS" in line.upper():
            break
        if line.strip():
            section_lines.append(line.strip())

    rows = []
    buffer_lines = []
    for line in section_lines:
        if "$" in line:
            if buffer_lines:
                merged_line = " ".join(buffer_lines + [line])
                rows.append(merged_line.strip())
                buffer_lines = []
            else:
                rows.append(line.strip())
        else:
            buffer_lines.append(line)
    if buffer_lines:
        rows.append(" ".join(buffer_lines))

    extracted = []
    for row in rows:
        if "Total State Standard Premium".upper() in row.upper():
            m_total = re.search(r'Total State Standard Premium\s*\$\s*([-]?\d[\d,\.]*)', row, re.IGNORECASE)
            if m_total:
                premium_val = "$" + m_total.group(1)
            else:
                premium_val = ""
            extracted.append(("", "Total State Standard Premium", premium_val))
            extracted.append(("", "", ""))
            continue

        pattern = re.compile(
            r'^(?:\D*?)(\d+)\s+(.*?)\s*\$\s*([-]?\d[\d,\.]+)(.*)$'
        )
        match = pattern.match(row)
        if match:
            code_no = match.group(1).strip()
            desc = match.group(2).strip()
            premium_val = "$" + match.group(3).replace(" ", "")
            leftover = match.group(4).strip()
            if leftover:
                desc += " " + leftover
            extracted.append((code_no, desc, premium_val))
        else:
            if "$" not in row:
                continue
            parts = row.split("$", 1)
            left = parts[0].strip()
            right = parts[1].strip()
            tokens = left.split(maxsplit=1)
            if len(tokens) == 2:
                code_no, desc = tokens[0], tokens[1]
            else:
                code_no, desc = "", left
            premium_val = "$" + right
            extracted.append((code_no, desc, premium_val))
    return extracted

#########################################
# NEW: Extract State Segments (Additional Tables)
#########################################
def extract_state_segments(lines):
    """
    Splits the pdfplumber lines into segments based on each occurrence of
    'SCHEDULE OF OPERATIONS'. This function does not alter any existing data
    mappings; it simply divides the entire pdfplumber text into chunks.
    """
    segments = []
    current_segment = []
    for line in lines:
        # Look for "SCHEDULE OF OPERATIONS" (case-insensitive)
        if "SCHEDULE OF OPERATIONS" in line.upper():
            if current_segment:
                segments.append(current_segment)
            current_segment = [line]
        else:
            if current_segment:
                current_segment.append(line)
    if current_segment:
        segments.append(current_segment)
    return segments

#########################################
# NEW: Helper function to parse a single line into columns
#########################################
def parse_line_into_columns(line: str) -> list:
    """
    Splits a single line into [Number, Edition, Description].
    1) Looks for a token matching a MM-YYYY pattern as 'Edition'.
    2) Everything before that is 'Number'; everything after is 'Description'.
    3) If no Edition is found, the entire line is treated as Number and Description is empty.
    """
    tokens = line.split()
    number_parts = []
    edition = ""
    description_parts = []

    # Pattern for something like 09-2008
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

#########################################
# NEW: Modified Policy Forms Extraction for Workers Comp
#########################################
def parse_policy_forms(text: str) -> dict:
    """
    1) Locate "SCHEDULE OF FORMS AND ENDORSEMENTS".
    2) From that point on, look for coverage type titles we want to parse.
    3) For each coverage title, skip the next line if it's "Number Edition Description".
    4) Parse each line using parse_line_into_columns():
       - If Edition is empty, treat it as a continuation of the previous row’s Description.
       - Otherwise, it starts a new row.
    5) Return a dict: { coverage_title: [ [Number, Edition, Description], ... ] }
    Updated to look for:
      - "Commercial Common Forms"
      - "Commercial Workers Compensation"
    """
    coverage_titles = [
        "Commercial Common Forms",
        "Commercial Workers Compensation"
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

        # If line is one of our coverage titles, start a new section
        if line in coverage_titles:
            current_title = line
            sections[current_title] = []
            i += 1
            # If the next line is "Number Edition Description", skip it
            if i < len(relevant_lines):
                possible_header_line = relevant_lines[i].lower().strip()
                if ("number" in possible_header_line and
                    "edition" in possible_header_line and
                    "description" in possible_header_line):
                    i += 1
            continue

        # If we are currently in a coverage section
        if current_title:
            # If a new "Commercial " heading appears that isn't in coverage_titles, end the current section
            if line.startswith("Commercial ") and (line not in coverage_titles):
                current_title = None
                i -= 1  # Reprocess this line next loop
            else:
                row = parse_line_into_columns(line)
                number, edition, description = row
                # If no edition => treat as a continuation
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

#########################################
# 8) Streamlit Main
#########################################
def main():
    st.title("PDF Policy Extraction (FEIN numeric fix)")

    uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])
    if uploaded_file is not None:
        file_bytes = uploaded_file.read()

        # pdfplumber lines
        try:
            text_pdfplumber = get_pdf_text_pdfplumber(file_bytes)
            lines_pdfplumber = text_pdfplumber.splitlines()
        except Exception as e:
            lines_pdfplumber = []
            st.sidebar.write(f"pdfplumber error: {e}")

        # PDFMiner lines
        lines_pdfminer = get_pdf_lines(file_bytes)

        # Debug method (your original debug options)
        debug_method = st.sidebar.selectbox(
            "Select Debug Extraction Method", 
            ["None", "PDFMiner", "PyPDF2", "pdfplumber", "Camelot Stream", "PyMuPDF"]
        )
        
        if debug_method == "PDFMiner":
            st.sidebar.markdown("### Debug: PDFMiner Extraction")
            for idx, ln in enumerate(lines_pdfminer):
                st.sidebar.write(f"{idx}: {ln}")
        elif debug_method == "PyPDF2":
            try:
                text_pypdf2 = get_pdf_text_pypdf2(file_bytes)
                lines_pypdf2 = text_pypdf2.splitlines()
                for idx, ln in enumerate(lines_pypdf2):
                    st.sidebar.write(f"{idx}: {ln}")
            except Exception as e:
                st.sidebar.write(f"PyPDF2 error: {e}")
        elif debug_method == "pdfplumber":
            st.sidebar.markdown("### Debug: pdfplumber Extraction")
            for idx, ln in enumerate(lines_pdfplumber):
                st.sidebar.write(f"{idx}: {ln}")
            st.sidebar.markdown("### Debug: pdfplumber Extracted Tables")
            import pdfplumber
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                for i, page in enumerate(pdf.pages):
                    table = page.extract_table()
                    if table:
                        st.sidebar.write(f"Page {i} Table:")
                        st.sidebar.write(table)
        elif debug_method == "Camelot Stream":
            try:
                import camelot
                tables = get_camelot_stream_tables(file_bytes)
                st.sidebar.markdown("### Debug: Camelot Stream Extraction (Raw Data)")
                for idx, table in enumerate(tables):
                    st.sidebar.write(f"Table {idx} raw data:")
                    st.sidebar.write(table.data)
            except Exception as e:
                st.sidebar.write(f"Camelot Stream error: {e}")
        elif debug_method == "PyMuPDF":
            try:
                text_pymupdf = get_pdf_text_pymupdf(file_bytes)
                lines_pymupdf = text_pymupdf.splitlines()
                st.sidebar.markdown("### Debug: PyMuPDF Extraction")
                for idx, ln in enumerate(lines_pymupdf):
                    st.sidebar.write(f"{idx}: {ln}")
            except Exception as e:
                st.sidebar.write(f"PyMuPDF error: {e}")

        # -----------------------
        # Table 1: Workers Comp Policy Information
        # -----------------------
        policy_info = extract_policy_information(lines_pdfminer)
        if not policy_info["Policy No."]:
            try:
                for i, line in enumerate(lines_pdfplumber):
                    if "States Government under the Act." in line:
                        for j in range(i, min(i + 4, len(lines_pdfplumber))):
                            match = re.search(r'Policy Number:\s*(.*)', lines_pdfplumber[j], re.IGNORECASE)
                            if match:
                                policy_info["Policy No."] = match.group(1).strip()
                                break
                        if policy_info["Policy No."]:
                            break
            except Exception:
                pass
        
        policy_table_data = [
            ("Date", policy_info["Date"]),
            ("Rating Company", policy_info["Rating Company"]),  # New second row added
            ("Quote No.", policy_info["Quote No."]),
            ("Policy No.", policy_info["Policy No."]),
            ("NCCI Carrier Code No.", policy_info["NCCI Carrier Code No."]),
            ("FEIN", policy_info["FEIN"]),
            ("Risk ID No.", policy_info["Risk ID No."]),
            ("Bureau File No.", policy_info["Bureau File No."]),
            ("Entity of Insured", policy_info["Entity of Insured"]),
            ("Proposed Policy Period", policy_info["Proposed Policy Period"]),
            ("Named Insured", policy_info["Named Insured"]),
            ("DBA", policy_info["DBA"]),
            ("Insured Address", policy_info["Insured Address"]),
            ("Insured City, State & Zip", policy_info["Insured City, State & Zip"]),
            ("Agent Name", policy_info["Agent Name"]),
            ("Agent Phone", policy_info["Agent Phone"]),
            ("Agent Address", policy_info["Agent Address"]),
            ("Agent City, State & Zip", policy_info["Agent City, State & Zip"])
        ]
        df_policy = pd.DataFrame(policy_table_data, columns=["Field", "Value"])
        st.markdown("### Policy Information")
        html_policy = df_policy.to_html(index=False)
        html_policy = make_table_cells_editable(html_policy)
        st.markdown(html_policy, unsafe_allow_html=True)

        # -----------------------
        # Table 2: Coverages
        # -----------------------
        coverages, premiums = extract_coverages(lines_pdfminer)
        if coverages:
            df_coverages = pd.DataFrame(list(zip(coverages, premiums)), columns=["Coverage", "Premium"])
            st.markdown("### Coverages")
            html_coverages = df_coverages.to_html(index=False)
            html_coverages = make_table_cells_editable(html_coverages)
            st.markdown(html_coverages, unsafe_allow_html=True)

        # -----------------------
        # Workers Comp Table
        # -----------------------
        workers_comp_rows = extract_workers_comp_table(lines_pdfplumber)
        if workers_comp_rows:
            df_workers = pd.DataFrame(workers_comp_rows, columns=["Coverage", "Limit", "Type"])
            st.markdown("<h3 style='text-align: left;'>Coverage</h3>", unsafe_allow_html=True)
            html_workers = df_workers.to_html(index=False)
            html_workers = make_table_cells_editable(html_workers)
            st.markdown(html_workers, unsafe_allow_html=True)
        else:
            st.info("No Workers Compensation and Employers Liability Quote Proposal information found in the PDF.")
        
        # -----------------------
        # Table 3: Additional Premium Info
        # -----------------------
        table3_rows = extract_table_3_pdfplumber(file_bytes)
        if table3_rows:
            df_t3 = pd.DataFrame(table3_rows, columns=["Description", "Premium"])
            st.markdown("### Additional Premium Info")
            html_t3 = df_t3.to_html(index=False)
            html_t3 = make_table_cells_editable(html_t3)
            st.markdown(html_t3, unsafe_allow_html=True)
        else:
            st.info("No Table 3 data found in the PDF.")

        # -----------------------
        # State-specific Schedule of Operations
        # -----------------------
        state_segments = extract_state_segments(lines_pdfplumber)
        st.markdown("## State-specific Schedule of Operations")

        if state_segments:
            for seg in state_segments:
                # Attempt to detect the state name
                state_name = "Unknown State"
                for i, txt in enumerate(seg):
                    if "SCHEDULE OF OPERATIONS" in txt.upper() and (i + 1) < len(seg):
                        candidate = seg[i+1].strip()
                        if candidate and "QUOTE NO" not in candidate.upper():
                            state_name = candidate
                        break

                # If we accidentally detect "EST ANNUAL" as a segment, skip it
                if state_name.upper() == "EST ANNUAL":
                    continue

                st.markdown(f"### {state_name}")

                # Schedule of Operations
                state_table_rows, state_table_subtotal = extract_schedule_operations_table(seg)
                if state_table_rows:
                    df_state_table = pd.DataFrame(
                        state_table_rows,
                        columns=[
                            "Loc",
                            "ST",
                            "Code No.",
                            "Classification",
                            "Premium Basis Total Estimated Annual Remuneration",
                            "Rate Per $100 of Remuneration",
                            "Estimated Annual Premium"
                        ]
                    )
                    html_state_table = df_state_table.to_html(index=False)
                    html_state_table = make_table_cells_editable(html_state_table)
                    st.markdown(html_state_table, unsafe_allow_html=True)

                # Subtotal
                if state_table_subtotal:
                    df_state_sub = pd.DataFrame([state_table_subtotal], columns=["Subtotal", "Description", "Amount"])
                    html_state_sub = df_state_sub.to_html(index=False)
                    html_state_sub = make_table_cells_editable(html_state_sub)
                    st.markdown(html_state_sub, unsafe_allow_html=True)

                # Additional Premium Info for that segment
                state_additional_premium = extract_additional_premium_info(seg)
                if state_additional_premium:
                    df_state_additional = pd.DataFrame(state_additional_premium, columns=["Code No.", "Description", "Premium"])
                    html_state_additional = df_state_additional.to_html(index=False)
                    html_state_additional = make_table_cells_editable(html_state_additional)
                    st.markdown(html_state_additional, unsafe_allow_html=True)
        else:
            st.info("No state-specific segments found in the PDF.")
        
        # -----------------------
        # NEW: Workers Compensation Forms Extraction
        # -----------------------
        st.markdown("## Workers Compensation Forms")
        # Use the pdfplumber text we extracted
        forms_sections = parse_policy_forms(text_pdfplumber)
        if forms_sections:
            for title, rows in forms_sections.items():
                st.markdown(f"### {title}")
                if rows:
                    df_forms = pd.DataFrame(rows, columns=["Number", "Edition", "Description"])
                    html_forms = df_forms.to_html(index=False)
                    html_forms = make_table_cells_editable(html_forms)
                    st.markdown(html_forms, unsafe_allow_html=True)
                else:
                    st.info(f"No rows found under {title}.")
        else:
            st.info("No Workers Compensation Forms sections found in the PDF.")
    else:
        st.info("Please upload a PDF file to begin.")

if __name__ == "__main__":
    main()
