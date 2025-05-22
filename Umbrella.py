import io
import re
import pandas as pd
import pdfplumber
import streamlit as st  # only needed if you are running this as a standalone app

def ensure_file_like(pdf_file):
    if isinstance(pdf_file, bytes):
        return io.BytesIO(pdf_file)
    return pdf_file

def make_table_cells_editable(html_str: str) -> str:
    """
    Insert contenteditable="true" into each <td> tag so the user can type/edit values in the browser.
    """
    pattern = r'(<td)([^>]*>)'
    replace = r'<td contenteditable="true"\2'
    return re.sub(pattern, replace, html_str, flags=re.IGNORECASE)

def parse_line_into_columns(line: str) -> list:
    """
    Splits a single line into [Number, Edition, Description].
    Looks for a token matching MM-YYYY as 'Edition'. If not found,
    entire line is 'Number' and Edition/Description are blank.
    """
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

def parse_policy_forms_for_umbrella(text: str) -> dict:
    """
    Parse forms starting at "SCHEDULE OF FORMS AND ENDORSEMENTS"
    for "Commercial Umbrella Coverage Part".
    """
    coverage_titles = ["Commercial Umbrella Coverage Part"]
    start_phrase = "SCHEDULE OF FORMS AND ENDORSEMENTS"
    start_index = text.find(start_phrase)
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
            # Skip header if present
            if i < len(relevant_lines):
                possible_header_line = relevant_lines[i].lower().strip()
                if ("number" in possible_header_line and
                    "edition" in possible_header_line and
                    "description" in possible_header_line):
                    i += 1
            continue
        if current_title:
            # If a new section starts, break out of the current umbrella section
            if line.startswith("Commercial ") and (line not in coverage_titles):
                current_title = None
                i -= 1  # reprocess this line
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

def extract_umbrella_data(pdf_file):
    """
    Extract umbrella-related data from the PDF file.
    Returns a dictionary with the following keys:
      - "CoveragePremium": DataFrame for Coverage & Premium table.
      - "Limits": DataFrame for Limits of Insurance table.
      - "Retention": DataFrame for Self-Insured Retention.
      - "Schedule": List of tuples (header, DataFrame) for Schedule of Underlying Insurance.
      - "PolicyForms": dict of DataFrames for Umbrella Policy Forms.
    """
    pdf_file = ensure_file_like(pdf_file)
    # Extract full text from the PDF.
    full_text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            full_text += page_text + "\n"
    # Split the full text into lines.
    lines = full_text.splitlines()

    # ---------------------------
    # Coverage & Premium Extraction
    # ---------------------------
    cp_start_keyword = "UMBRELLA OR EXCESS LIABILITY COVERAGES PREMIUM"
    cp_stop_keyword = "LIMITS OF INSURANCE"
    found_cp_start = False
    cp_lines = []
    for line in lines:
        stripped = line.strip()
        if not found_cp_start:
            if cp_start_keyword in stripped.upper():
                found_cp_start = True
        else:
            if cp_stop_keyword in stripped.upper():
                break
            cp_lines.append(stripped)
    cp_rows = []
    for ln in cp_lines:
        if ln.upper().startswith("TOTAL QUOTE PREMIUM") and ln == ln.upper():
            continue
        if "$" in ln:
            parts = ln.rsplit("$", 1)
            if len(parts) == 2:
                coverage_text = parts[0].strip()
                premium_text = "$" + parts[1].strip()
                cp_rows.append({"Coverage": coverage_text, "Premium": premium_text})
    coverage_premium_data = pd.DataFrame(cp_rows) if cp_rows else pd.DataFrame()

    # ---------------------------
    # Limits of Insurance Extraction
    # ---------------------------
    # We first look for "COMMERCIAL LIABILITY UMBRELLA QUOTE PROPOSAL",
    # then from there find "LIMITS OF INSURANCE".
    umbrella_keyword = "COMMERCIAL LIABILITY UMBRELLA QUOTE PROPOSAL"
    li_start_keyword = "LIMITS OF INSURANCE"
    found_umbrella_keyword = False
    found_li_start = False
    li_lines = []

    for line in lines:
        stripped = line.strip().upper()

        if not found_umbrella_keyword:
            # Search for "COMMERCIAL LIABILITY UMBRELLA QUOTE PROPOSAL"
            if umbrella_keyword in stripped:
                found_umbrella_keyword = True
            continue

        if found_umbrella_keyword and not found_li_start:
            # Next, we look for "LIMITS OF INSURANCE"
            if li_start_keyword in stripped:
                found_li_start = True
            continue

        # Once both are found, start capturing lines
        if found_umbrella_keyword and found_li_start:
            # Stop if we reach Self-Insured Retention
            if "SELF-INSURED RETENTION" in stripped:
                break
            # Also stop if we find an all-caps line (a new heading) without a dollar sign
            if stripped and stripped == stripped.upper() and "$" not in stripped:
                break
            li_lines.append(line.strip())

    li_rows = []
    for ln in li_lines:
        if not ln or ln.startswith("("):
            continue
        if "$" in ln:
            parts = ln.rsplit("$", 1)
            if len(parts) == 2:
                coverage_text = re.sub(r'\.+', '', parts[0]).strip()
                tokens = parts[1].strip().split()
                if tokens:
                    limit_token = tokens[0]
                    limit_text = "$" + limit_token
                    li_rows.append({"Coverage": coverage_text, "Limits": limit_text})

    limits_data = pd.DataFrame(li_rows) if li_rows else pd.DataFrame()

    # ---------------------------
    # Self-Insured Retention Extraction
    # ---------------------------
    retention_line = None
    for line in lines:
        if "SELF-INSURED RETENTION:" in line.upper():
            retention_line = line.strip()
            break
    if retention_line:
        retention_line = re.sub(r"^\d+\.\s*", "", retention_line)
        parts = retention_line.split(":", 1)
        if len(parts) == 2:
            retention_label = parts[0].strip()
            retention_value = parts[1].strip()
        else:
            retention_label = "SELF-INSURED RETENTION"
            retention_value = ""
        retention_data = pd.DataFrame([{"Type": retention_label, "Data": retention_value}])
    else:
        retention_data = pd.DataFrame()

    # ---------------------------
    # Schedule of Underlying Insurance Extraction
    # ---------------------------
    schedule_start_keyword = "SCHEDULE OF UNDERLYING INSURANCE"
    termination_markers = ["COMMERCIAL INLAND MARINE", "QUOTE PROPOSAL"]
    found_schedule = False
    schedule_lines = []
    for line in lines:
        stripped = line.strip()
        if not found_schedule:
            if schedule_start_keyword in stripped.upper():
                found_schedule = True
        else:
            term_found = False
            for marker in termination_markers:
                if marker in stripped.upper():
                    term_found = True
                    break
            if term_found:
                break
            schedule_lines.append(stripped)

    groups = []
    current_group = None
    for ln in schedule_lines:
        if not ln:
            continue
        if (":" not in ln) and ("$" not in ln):
            if current_group is not None:
                groups.append(current_group)
            current_group = {"Header": ln, "Rows": []}
        else:
            if ":" in ln:
                parts = ln.split(":", 1)
                key = parts[0].strip()
                value = parts[1].strip()
            elif "$" in ln:
                parts = ln.rsplit("$", 1)
                key = parts[0].strip()
                value = "$" + parts[1].strip()
            else:
                key = ln
                value = ""
            if current_group is not None:
                current_group["Rows"].append({"Type": key, "Data": value})
    if current_group is not None:
        groups.append(current_group)

    schedule_tables = []
    for group in groups:
        header = group["Header"]
        rows = group["Rows"]
        if rows:
            df = pd.DataFrame(rows)
            schedule_tables.append((header, df))

    # ---------------------------
    # Policy Forms Extraction (Commercial Umbrella)
    # ---------------------------
    policy_forms_sections = parse_policy_forms_for_umbrella(full_text)
    policy_forms = {}
    for title, rows in policy_forms_sections.items():
        policy_forms[title] = pd.DataFrame(rows, columns=["Number", "Edition", "Description"])

    return {
        "CoveragePremium": coverage_premium_data,
        "Limits": limits_data,
        "Retention": retention_data,
        "Schedule": schedule_tables,
        "PolicyForms": policy_forms,
    }

if __name__ == "__main__":
    st.title("Umbrella Data Extractor")
    uploaded_file = st.file_uploader("Drag and drop your PDF file here", type=["pdf"])
    if uploaded_file is not None:
        data = extract_umbrella_data(uploaded_file.read())
        st.write("## COMMERCIAL LIABILITY UMBRELLA")
        
        # Coverage & Premium Table
        st.write("### Coverage & Premium Table")
        cp = data.get("CoveragePremium")
        if cp is not None and not cp.empty:
            html_cp = make_table_cells_editable(cp.to_html(index=False))
            st.markdown(html_cp, unsafe_allow_html=True)
        else:
            st.write("No Coverage & Premium data found.")
        
        # Limits of Insurance Table
        st.write("### Limits of Insurance Table")
        li = data.get("Limits")
        if li is not None and not li.empty:
            html_li = make_table_cells_editable(li.to_html(index=False))
            st.markdown(html_li, unsafe_allow_html=True)
        else:
            st.write("No Limits of Insurance data found.")
        
        # SELF-INSURED RETENTION
        st.write("### SELF-INSURED RETENTION")
        ret = data.get("Retention")
        if ret is not None and not ret.empty:
            html_ret = make_table_cells_editable(ret.to_html(index=False))
            st.markdown(html_ret, unsafe_allow_html=True)
        else:
            st.write("No SELF-INSURED RETENTION data found.")
        
        # SCHEDULE OF UNDERLYING INSURANCE
        st.write("### SCHEDULE OF UNDERLYING INSURANCE")
        schedule = data.get("Schedule")
        if schedule:
            for header, df in schedule:
                st.write(f"#### {header}")
                html_sched = make_table_cells_editable(df.to_html(index=False))
                st.markdown(html_sched, unsafe_allow_html=True)
        else:
            st.write("No Schedule of Underlying Insurance data found.")
        
        # Umbrella Policy Forms
        st.write("### Umbrella Policy Forms")
        pf = data.get("PolicyForms")
        if pf:
            for title, df in pf.items():
                st.write(f"#### {title}")
                html_pf = make_table_cells_editable(df.to_html(index=False))
                st.markdown(html_pf, unsafe_allow_html=True)
        else:
            st.write("No Umbrella Policy Forms data found.")
