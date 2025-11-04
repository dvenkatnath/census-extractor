"""
LLM-based Extraction Engine
============================
Robust extraction using full data context (nf7.py approach, but modular).

This module sends ALL data to LLM for extraction, providing:
- Full context understanding
- Adaptive to any file structure
- Handles name splitting, dependent grouping, relationship identification
- Returns standardized canonical format
"""

import pandas as pd
import json
import math
import time
import re
import os
import tempfile
from groq import Groq
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configuration
api_key = os.getenv("GROQ_API_KEY")
if not api_key:
    raise ValueError("❌ GROQ_API_KEY not found in environment variables. Please set it in .env file.")

client = Groq(api_key=api_key)
LLM_MODEL = "llama-3.3-70b-versatile"
MAX_INPUT_CHARS = 40000
MAX_OUTPUT_CHECK = 40000

# Field mapping from nf7.py format to canonical format
FIELD_MAPPING = {
    "last_name": "Last Name",
    "first_name": "First Name",
    "employee_name": "Employee Name",
    "home_zip_code": "Home Zip Code",
    "dob": "DOB",
    "gender": "Gender",
    "medical_coverage": "Medical Coverage",
    "medical_coverage_level": "Medical Coverage Level",
    "vision_coverage": "Vision Coverage",
    "vision_coverage_level": "Vision Coverage Level",
    "dental_coverage": "Dental Coverage",
    "dental_coverage_level": "Dental Coverage Level",
    "cobra_participation": "COBRA Participation (Y/N)",
    "relationship_to_employee": "Relationship To employee",
    "dependent_of_employee_row": "Dependent Of Employee Row"
}


def clean_json_output(output_text):
    """Extract and fix valid JSON content from messy LLM output."""
    if not output_text:
        return "[]"

    text = output_text.strip()
    text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.MULTILINE).strip()

    try:
        json.loads(text)
        return text
    except:
        pass

    matches = re.findall(r'(\[.*?\]|\{.*?\})', text, re.DOTALL)
    if not matches:
        return "[]"

    valid_parts = []
    for m in matches:
        json_str = re.sub(r"'", '"', m)
        json_str = re.sub(r'([{,]\s*)([A-Za-z0-9_]+)\s*:', r'\1"\2":', json_str)
        json_str = re.sub(r",\s*([\]}])", r"\1", json_str)

        try:
            obj = json.loads(json_str)
            if isinstance(obj, list):
                valid_parts.extend(obj)
            else:
                valid_parts.append(obj)
        except:
            continue

    if valid_parts:
        return json.dumps(valid_parts, ensure_ascii=False, indent=2)
    return "[]"


def get_full_llm_output(prompt, model=LLM_MODEL, log=None):
    """Send prompt to Groq and request continuation if truncated."""
    full_output = ""
    part = 1

    while True:
        if log:
            log(f"🧠 Sending LLM request part {part}...")

        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=16384,
            temperature=0.2
        )

        output = response.choices[0].message.content.strip()
        full_output += output

        if len(output) >= MAX_OUTPUT_CHECK and not output.strip().endswith(("]", "}")):
            if log:
                log(f"⏩ Output truncated. Requesting continuation (part {part+1})...")
            prompt = "Continue from where you left off. Do not repeat earlier content."
            part += 1
            time.sleep(1)
            continue
        break

    return full_output


def read_all_sheets(file_path_or_object, log=None):
    """
    Read all Excel sheets into a single combined text representation.
    
    Args:
        file_path_or_object: Either a file path (str) or file-like object (for Streamlit uploads)
        log: Optional logging function
    
    Returns:
        str: Combined text representation of all sheets
    """
    # Handle file objects (Streamlit uploads)
    if hasattr(file_path_or_object, 'read'):
        # Save to temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(file_path_or_object.read())
            tmp_path = tmp_file.name
        file_path = tmp_path
        cleanup_temp = True
    else:
        file_path = file_path_or_object
        cleanup_temp = False
    
    try:
        xls = pd.ExcelFile(file_path)
        combined_text = ""

        if log:
            log(f"📘 Loaded workbook with {len(xls.sheet_names)} sheets")

        for sheet_name in xls.sheet_names:
            # Skip irrelevant sheets early
            if sheet_name.lower() in ["company info", "enrollment info"]:
                if log:
                    log(f"⚠️ Skipping irrelevant sheet '{sheet_name}'")
                continue

            df = xls.parse(sheet_name)
            if df.empty:
                continue

            # Skip if no employee-related content
            combined_cols = " ".join(df.columns.astype(str))
            if not any(k in combined_cols.lower() for k in ["name", "dob", "employee", "dependent", "zip"]):
                continue

            if log:
                log(f"📄 Including sheet '{sheet_name}' ({len(df)} rows)")

            combined_text += f"\n\n### SHEET: {sheet_name}\n"
            combined_text += df.to_string(index=False)

        if not combined_text.strip():
            raise ValueError("❌ No valid employee-related data found in any sheet.")

        return combined_text
    finally:
        if cleanup_temp and os.path.exists(file_path):
            os.unlink(file_path)


def convert_to_canonical_format(records):
    """
    Convert nf7.py format to canonical format used by app.py.
    
    Args:
        records: List of dicts in nf7.py format
    
    Returns:
        List of dicts in canonical format
    """
    canonical_records = []
    
    for record in records:
        canonical_record = {}
        
        # Map fields
        for nf7_field, canonical_field in FIELD_MAPPING.items():
            value = record.get(nf7_field, "")
            # Convert None, "nan", etc. to empty string
            if value is None or str(value).lower() in ["nan", "none", "null", ""]:
                canonical_record[canonical_field] = ""
            else:
                canonical_record[canonical_field] = str(value).strip()
        
        # Handle Dependent (Y/N) field
        relationship = str(record.get("relationship_to_employee", "")).lower().strip()
        if relationship in ["spouse", "child", "son", "daughter", "wife", "husband", "dependent"]:
            canonical_record["Dependent (Y/N)"] = "Y"
        elif relationship in ["employee", "self", ""]:
            canonical_record["Dependent (Y/N)"] = "N"
        else:
            # Determine based on relationship
            canonical_record["Dependent (Y/N)"] = "Y" if relationship else "N"
        
        # Add missing canonical fields with empty defaults
        canonical_fields = [
            "First Name", "Last Name", "Employee Name", "DOB", "Gender",
            "Relationship To employee", "Dependent (Y/N)", "Medical Coverage",
            "Medical Plan Name", "Dental Coverage", "Dental Plan Name",
            "Vision Coverage", "Vision Plan Name", "COBRA Participation (Y/N)"
        ]
        
        for field in canonical_fields:
            if field not in canonical_record:
                canonical_record[field] = ""
        
        canonical_records.append(canonical_record)
    
    return canonical_records


def extract_with_full_context(file_path_or_object, log=None):
    """
    Extract census data using full LLM context (robust approach).
    
    This is the main entry point for full-context extraction.
    It combines all sheets, chunks if needed, and sends to LLM.
    
    Args:
        file_path_or_object: Either a file path (str) or file-like object
        log: Optional logging function (for Streamlit integration)
    
    Returns:
        pandas.DataFrame: Extracted data in canonical format
    """
    all_results = []

    try:
        # Step 1: Combine all sheets
        combined_text = read_all_sheets(file_path_or_object, log=log)
        total_len = len(combined_text)
        num_chunks = math.ceil(total_len / MAX_INPUT_CHARS)

        if log:
            log(f"🧩 Combined all sheets into {num_chunks} chunk(s) (total {total_len:,} chars)")

        # Step 2: Process each chunk
        for i in range(num_chunks):
            chunk_text = combined_text[i * MAX_INPUT_CHARS : (i + 1) * MAX_INPUT_CHARS]

            prompt = f"""
            You are an expert data extractor.
            From the following combined census workbook text (chunk {i+1}/{num_chunks}),
            extract all employee and dependent records in sequence.
            Return only valid JSON list — no explanations, no markdown.

            Fields:
            ["last_name","first_name","employee_name","home_zip_code",
             "dob","gender","medical_coverage","medical_coverage_level",
             "vision_coverage","vision_coverage_level",
             "dental_coverage","dental_coverage_level",
             "cobra_participation","relationship_to_employee","dependent_of_employee_row"]

            Rules:
                - Read all worksheets, find employee info, and fetch data once found do not process other sheets.
                - Read all the rows in the input file, in the excel worksheet if empty rows found skip and move to next row and read till the end of row with data.
                - Use the exact field names provided.
                - If COBRA Participant is missing, return "No".
                - If First Name is missing, skip that row.
                - Dependents: split 'DEPENDENT n NAME' into First Name and Last Name.
                - Relationship to Employee: use exactly as shown in Excel (no grouping by last name).
                - Employee rows: Relationship to Employee = "Employee", Dependent of Employee Row = null.
                - Dependents: Dependent of Employee Row should link to the Employee in order of appearance (Slno).
                - Dependent data in separate row, not as sub node in json
                - for medical coverage, dental coverage and vision coverage match and pick the plan name not provider
                - Avoid processing sheet named "company info" and "enrollment info"
                - Read all the sheets data keep it in memory, find the employee data and dependent data and group it. provide the output in meaningful family sequence
                - Avoid timestamp on Date of Birth column, only provide available data
                - Compensation type is not coverage level. Do not map
                - do not return NAN, return blank
                - If dependent fields in another sheet, map it with the employee data based the first name and last name
                - do not map worker compensation code and salary to coverage and coverage level
                - if dependent DATA found in another sheet of a excel, map THE DEPENDENT WITH THE EMPLOYEE LAST name and generate the OUTPUT WITH dependent row next to employee mapped
                - if relationship_to_employee not found leave blank
                - If employee data repeated in another sheet, do not repeat in output, employee row is unique also provide the dependent name in name column not the employee
                - provide employee name in respective dependent_of_employee_row
                - in coverage level, do not fill data split from coverage. if coverage level not available provide blank

            Table chunk:
            {chunk_text}
            """

            output_text = get_full_llm_output(prompt, log=log)
            cleaned = clean_json_output(output_text)

            try:
                json_data = json.loads(cleaned)
                if isinstance(json_data, list):
                    all_results.extend(json_data)
                else:
                    all_results.append(json_data)
            except json.JSONDecodeError:
                if log:
                    log(f"⚠️ Could not parse JSON for chunk {i+1}")
                continue

            if log:
                log(f"✅ Finished chunk {i+1}/{num_chunks} ({len(all_results)} records so far)")
            time.sleep(1)  # Rate limiting

        if log:
            log(f"🎯 Total extracted records: {len(all_results)}")
        
        # Convert to canonical format
        canonical_records = convert_to_canonical_format(all_results)
        
        # Convert to DataFrame
        if canonical_records:
            df = pd.DataFrame(canonical_records)
            
            # Reorder columns: Move Relationship To employee, Dependent Of Employee Row, and Dependent (Y/N) next to Gender
            cols = df.columns.tolist()
            
            # Fields to move next to Gender
            relationship_fields = ["Relationship To employee", "Dependent Of Employee Row", "Dependent (Y/N)"]
            
            # Find Gender position
            if "Gender" in cols:
                gender_idx = cols.index("Gender")
                
                # Remove relationship fields from their current positions
                for field in relationship_fields:
                    if field in cols:
                        cols.remove(field)
                
                # Re-find Gender position after removals
                gender_idx = cols.index("Gender")
                
                # Insert relationship fields after Gender
                insert_pos = gender_idx + 1
                for field in relationship_fields:
                    if field in df.columns:
                        cols.insert(insert_pos, field)
                        insert_pos += 1
                
                df = df[cols]
            
            return df
        else:
            # Return empty DataFrame with canonical columns
            canonical_fields = [
                "First Name", "Last Name", "Employee Name", "DOB", "Gender",
                "Relationship To employee", "Dependent (Y/N)", "Medical Coverage",
                "Medical Plan Name", "Dental Coverage", "Dental Plan Name",
                "Vision Coverage", "Vision Plan Name", "COBRA Participation (Y/N)"
            ]
            return pd.DataFrame(columns=canonical_fields)
            
    except Exception as e:
        if log:
            log(f"❌ Error during extraction: {e}")
        # Return empty DataFrame on error (with proper column order)
        canonical_fields = [
            "First Name", "Last Name", "Employee Name", "DOB", "Gender",
            "Relationship To employee", "Dependent Of Employee Row", "Dependent (Y/N)",
            "Medical Coverage", "Medical Plan Name", "Dental Coverage", "Dental Plan Name",
            "Vision Coverage", "Vision Plan Name", "COBRA Participation (Y/N)"
        ]
        return pd.DataFrame(columns=canonical_fields)


def produce_stats_for_llm(all_sheets, extracted_df):
    """Produce statistics for LLM extraction mode in markdown table format."""
    import io
    out = io.StringIO()
    
    # Markdown format
    out.write("## Extraction Statistics\n\n")
    
    # Summary table
    out.write("### Summary\n\n")
    out.write("| Metric | Value |\n")
    out.write("|--------|-------|\n")
    out.write(f"| **Filename (Excel)** | Uploaded file |\n")
    out.write(f"| **Total Sheets** | {len(all_sheets)} |\n")
    out.write(f"| **Extraction Method** | Full LLM Context Extraction |\n")
    out.write(f"| **Total Extracted Records** | {len(extracted_df)} |\n")
    
    # Count employees and dependents if available
    if 'Relationship To employee' in extracted_df.columns:
        employees = len(extracted_df[extracted_df['Relationship To employee'].str.lower().str.strip() == 'employee'])
        dependents = len(extracted_df) - employees
        out.write(f"| **Employees** | {employees} |\n")
        out.write(f"| **Dependents** | {dependents} |\n")
    
    out.write("\n### Sheet Details\n\n")
    out.write("| Sheet Name | Rows | Columns |\n")
    out.write("|------------|------|---------|\n")
    for name, df in all_sheets.items():
        out.write(f"| {name} | {df.shape[0]} | {df.shape[1]} |\n")
    
    return out.getvalue()

