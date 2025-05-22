import streamlit as st
import re
import os
import tempfile
from pdfminer.high_level import extract_text

def merge_header_lines(lines):
    """
    Merge consecutive lines if they form a known heading.
    For example, if the heading is split as:
      EMPLOYMENT-RELATED PRACTICES LIABILITY
      QUOTE PROPOSAL
    then merge them into:
      EMPLOYMENT-RELATED PRACTICES LIABILITY QUOTE PROPOSAL
    """
    merged = []
    i = 0
    while i < len(lines):
        if i < len(lines) - 1:
            line_lower = lines[i].lower()
            next_lower = lines[i+1].lower()
            if ("employment-related practices liability" in line_lower 
                and next_lower == "quote proposal"):
                merged.append("EMPLOYMENT-RELATED PRACTICES LIABILITY QUOTE PROPOSAL")
                i += 2
                continue
            if ("business auto" in line_lower 
                and next_lower == "quote proposal"):
                merged.append("BUSINESS AUTO QUOTE PROPOSAL")
                i += 2
                continue
        merged.append(lines[i])
        i += 1
    return merged

def parse_erp_quote_proposal(raw_text):
    """
    1) Merge lines so that "EMPLOYMENT-RELATED PRACTICES LIABILITY" and "QUOTE PROPOSAL"
       become one line.
    2) Identify the section starting at "EMPLOYMENT-RELATED PRACTICES LIABILITY QUOTE PROPOSAL"
       and ending before "BUSINESS AUTO QUOTE PROPOSAL".
    3) Extract the following fields from that chunk:
         - Aggregate Limit (the next line containing '$')
         - Each "Claim" Limit (the next line containing '$')
         - Deductible (the next line containing '$')
         - Retroactive Date (the immediate next line)
         - Estimated Total Premium (check the same line for a '$' or take the next line containing '$')
    """
    # Split the text into non-empty, stripped lines
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    merged_lines = merge_header_lines(lines)

    # Locate the start of the ERP section
    start_index = None
    for i, line in enumerate(merged_lines):
        if re.search(r"(?i)employment-related practices liability quote proposal", line):
            start_index = i
            break
    if start_index is None:
        return {}

    # Locate the end of the ERP section using "BUSINESS AUTO QUOTE PROPOSAL"
    end_index = None
    for j in range(start_index + 1, len(merged_lines)):
        if re.search(r"(?i)business auto quote proposal", merged_lines[j]):
            end_index = j
            break
    if end_index is None:
        end_index = len(merged_lines)

    chunk = merged_lines[start_index:end_index]

    def find_next_dollar_line(chunk_lines, start_idx):
        for x in range(start_idx + 1, len(chunk_lines)):
            if "$" in chunk_lines[x]:
                return chunk_lines[x].strip()
        return ""

    def dollar_in_same_line(text):
        match = re.search(r"\$\s*\d[\d,\.]*", text)
        return match.group(0).strip() if match else ""

    results = {
        "agg_limit_value": "",
        "each_claim_limit_value": "",
        "deductible_value": "",
        "retro_date": "",
        "est_premium": ""
    }

    for idx, line in enumerate(chunk):
        if re.match(r"(?i)^aggregate limit$", line):
            results["agg_limit_value"] = find_next_dollar_line(chunk, idx)
        elif re.match(r'(?i)^each\s*"claim"\s*limit$', line):
            results["each_claim_limit_value"] = find_next_dollar_line(chunk, idx)
        elif re.match(r"(?i)^deductible:?", line):
            results["deductible_value"] = find_next_dollar_line(chunk, idx)
        elif re.match(r"(?i)^retroactive date:?", line):
            if idx + 1 < len(chunk):
                results["retro_date"] = chunk[idx + 1].strip()
        elif re.match(r"(?i)^estimated total premium:?", line):
            same_line_dollar = dollar_in_same_line(line)
            if same_line_dollar:
                results["est_premium"] = same_line_dollar
            else:
                results["est_premium"] = find_next_dollar_line(chunk, idx)

    # Remove extra spaces
    for key in results:
        results[key] = results[key].replace(" ", "")

    # Format the premium to remove trailing zeros, e.g. "$2,008.00" → "$2,008"
    est = results["est_premium"]
    if est.startswith("$"):
        try:
            val = float(est.replace("$", "").replace(",", ""))
            if val.is_integer():
                results["est_premium"] = "$" + format(int(val), ",")
            else:
                formatted = f"{val:,.2f}"
                if formatted.endswith(".00"):
                    formatted = formatted[:-3]
                results["est_premium"] = "$" + formatted
        except:
            pass

    return results

def generate_html_table(parsed):
    """
    Create an HTML table with:
      - A caption that serves as the title ("EMPLOYMENT-RELATED PRACTICES LIABILITY").
      - A header row with a teal background and white text.
      - A single data row with a white background.
      - Each cell is editable.
    """
    teal_color = "#2D5D77"
    html = f"""
    <table style="border-collapse: collapse; width: 100%; margin-top: 20px;">
      <caption style="caption-side: top; text-align: center; font-size: 1.5em; font-weight: bold; margin-bottom: 10px;">
        EMPLOYMENT-RELATED PRACTICES LIABILITY
      </caption>
      <thead style="background-color: {teal_color}; color: white;">
        <tr>
          <th style="border: 1px solid #ccc; padding: 8px 12px;">Aggregate Limit</th>
          <th style="border: 1px solid #ccc; padding: 8px 12px;">Each 'Claim' Limit</th>
          <th style="border: 1px solid #ccc; padding: 8px 12px;">Deductible</th>
          <th style="border: 1px solid #ccc; padding: 8px 12px;">Retroactive Date</th>
          <th style="border: 1px solid #ccc; padding: 8px 12px;">Estimated Total Premium</th>
        </tr>
      </thead>
      <tbody>
        <tr style="background-color: white;">
          <td contenteditable="true" style="border: 1px solid #ccc; padding: 8px 12px;">{parsed.get("agg_limit_value", "")}</td>
          <td contenteditable="true" style="border: 1px solid #ccc; padding: 8px 12px;">{parsed.get("each_claim_limit_value", "")}</td>
          <td contenteditable="true" style="border: 1px solid #ccc; padding: 8px 12px;">{parsed.get("deductible_value", "")}</td>
          <td contenteditable="true" style="border: 1px solid #ccc; padding: 8px 12px;">{parsed.get("retro_date", "")}</td>
          <td contenteditable="true" style="border: 1px solid #ccc; padding: 8px 12px;">{parsed.get("est_premium", "")}</td>
        </tr>
      </tbody>
    </table>
    """
    return html

def main():
    st.title("PDF Data Extractor (HTML Table Output with Editable Cells)")

    uploaded_file = st.file_uploader("Drag and drop a PDF file", type=["pdf"])
    if uploaded_file is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.read())
            tmp.flush()
            tmp_path = tmp.name

        raw_text = extract_text(tmp_path)
        os.remove(tmp_path)

        parsed_values = parse_erp_quote_proposal(raw_text)
        if not any(parsed_values.values()):
            st.write("No data found for EMPLOYMENT-RELATED PRACTICES LIABILITY QUOTE PROPOSAL in this PDF.")
        else:
            html_table = generate_html_table(parsed_values)
            st.markdown(html_table, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
