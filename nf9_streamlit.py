import pandas as pd
import json
import math
import time
import re
import csv
from groq import Groq
import os
import streamlit as st
import io
import tempfile
from datetime import datetime

# =========================================================
# CONFIGURATION
# =========================================================
# Load API key from environment (required for Streamlit Cloud)
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
if not GROQ_API_KEY:
    st.error("❌ GROQ_API_KEY not found. Please set it in Streamlit Cloud Secrets.")
    st.stop()
client = Groq(api_key=GROQ_API_KEY)
LLM_MODEL = "openai/gpt-oss-120b"
MAX_INPUT_CHARS = 80000
MAX_OUTPUT_CHECK = 80000

# Streamlit page config
st.set_page_config(page_title="Census Data Extraction from Excel using LLM", layout="wide")

# Hide Streamlit default UI elements
hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# =========================================================
# LANDING PAGE WARNING
# =========================================================
if "warning_accepted" not in st.session_state:
    st.session_state.warning_accepted = False

if not st.session_state.warning_accepted:
    st.title("📊 Census Data Extraction from Excel using LLM")
    st.markdown("---")
    
    st.warning("""
    **⚠️ Caution:** This platform is not intended for sensitive data. Please remove all PII (Personally Identifiable Information) before uploading. 
    Any uploaded content may be processed by an AI model hosted in the cloud. You are solely responsible for ensuring compliance with data privacy and security regulations.
    """)
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if st.button("✅ I Understand and Agree to Proceed", type="primary", use_container_width=True):
            st.session_state.warning_accepted = True
            st.rerun()
    
    # Footer
    st.markdown("---")
    footer_html = """
    <div style="display: flex; justify-content: space-between; padding: 10px; font-size: 12px; color: #666;">
        <div>ConceptVine : Proof of Concept</div>
        <div>Powered by Shepardtri</div>
    </div>
    """
    st.markdown(footer_html, unsafe_allow_html=True)
    
    st.stop()

# =========================================================
# HELPER FUNCTIONS (UNCHANGED FROM nf9.py)
# =========================================================

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
       
        json_str=m
        json_str = re.sub(r"'", '"', json_str)
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
            time.sleep(2)
            continue
        break

    return full_output


def read_all_sheets(file_path, log=None):
    """Read all Excel sheets into a single combined text representation."""
    xls = pd.ExcelFile(file_path)
    combined_text = ""

    if log:
        log(f"📘 Loaded workbook '{file_path}' with {len(xls.sheet_names)} sheets")

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


def process_combined_excel(file_path, log=None):
    """Combine all sheets, chunk, and send to LLM sequentially."""
    all_results = []

    try:
        # Step 1: Combine all sheets
        combined_text = read_all_sheets(file_path, log=log)
        total_len = len(combined_text)
        num_chunks = math.ceil(total_len / MAX_INPUT_CHARS)

        if log:
            log(f"🧩 Combined all sheets into {num_chunks} chunks (total {total_len:,} chars)")

        # Step 2: Process each chunk
        for i in range(num_chunks):
            chunk_text = combined_text[i * MAX_INPUT_CHARS : (i + 1) * MAX_INPUT_CHARS]

            prompt = f"""
            You are an expert table data extractor.
            From the following combined census workbook text (chunk {i+1}/{num_chunks}),
            extract all employee census data from table and dependent records.
            do not skip any data from input 
            Return only valid JSON list — no explanations, no markdown.

            Fields:
            ["last_name","first_name","employee_name","home_zip_code",
             "dob","gender","medical_coverage","medical_coverage_level",
             "vision_coverage","vision_coverage_level",
             "dental_coverage","dental_coverage_level",
             "cobra_participation","relationship_to_employee","dependent_of_employee_row"]

            Rules:
                - Read all worksheets, find employee info, and fetch data once found do not process other sheets .
                - Read all the rows in the input file ,in the excel work sheet if empty rows found skip and move to next row and read till the end of row with data.
                - Use the exact field names provided.
                - in addition to field list provide running serial number for each row
                - If COBRA Participant is missing, return "No".
                - If First Name is missing, skip that row.
                - Dependents: split 'DEPENDENT n NAME' into First Name and Last Name.
                - Relationship to Employee: use exactly as shown in Excel (no grouping by last name).
                - Employee rows: Relationship to Employee = "Employee", Dependent of Employee Row = null.
                - Dependents: Dependent of Employee Row should link to the Employee in order of appearance (Slno).
                - Dependent data in separate row , not as sub node in json
                - for medical coverage , dental coverage and vision coverage match and pick the plan name not provider
                - Avoid processing sheet named "company info" and "enrollment info"
                - Read all the sheets data keep it in memory , find the employee data and dependent data and group it.provide the output in meaning full family sequence
                - Avoid timestamp on Date of Birth column , only provide available data
                - Compensation type is not coverage level.Do not map
                - do not retund NAN , return blank
                - If dependent fields in another sheet , map it with the employee data based the first name and last name
                - do not map worker compensation code and salary to coverage and coverage level
                - if dependent DATA found in another sheet of a excel, map THE DEPENDENT WITH THE EMPLOYEE lAST name and generate the OUTPUT WITH  dependent row next to employee mapped
                - if relationship_to_employee not found leave blank , do not map First Name (Subscriber / Dependent) 
                - If employee data repeated in another sheet , do not reapet in output , employee row is unique also provide the dependent name in name column not the employee
                - provide employee name in respective dependent_of_employee_row , do not row number
                - Coverage levels: Use the EXACT value from the Excel file. Valid values include abbreviations like "EE" or "E" (Employee), "F" (Family), "SP" (Spouse), "E+F" (Employee + Family), or descriptive text like "Employee Only", "Family", "Employee + Spouse". DO NOT use numbers, plan codes, deductibles (like "1500", "1000", etc.), or any numeric values. If coverage level column contains only numbers or is not explicitly provided, leave blank. Do not infer coverage level from plan names or split data from coverage field.
           

            Table chunk:
            {chunk_text}
            """

            output_text = get_full_llm_output(prompt, log=log)
            time.sleep(2)
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
            time.sleep(2)

        if log:
            log(f"🎯 Total extracted records: {len(all_results)}")
        return all_results

    except Exception as e:
        if log:
            log(f"❌ Error: {e}")
        return []


def json_to_csv_by_input_name(json_filepath):
    """Convert JSON (list of dicts) to CSV file with same name."""
    try:
        json_dir = os.path.dirname(json_filepath)
        base_name = os.path.splitext(os.path.basename(json_filepath))[0]
        csv_filepath = os.path.join(json_dir, base_name + ".csv")

        with open(json_filepath, "r", encoding="utf-8") as jf:
            data = json.load(jf)

        if not isinstance(data, list) or not data:
            print("❌ JSON file must contain a non-empty list.")
            return

        fieldnames = list(data[0].keys())
        with open(csv_filepath, "w", newline="", encoding="utf-8") as cf:
            writer = csv.DictWriter(cf, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(data)

        print(f"✅ CSV saved: {csv_filepath}")
    except Exception as e:
        print(f"❌ CSV conversion failed: {e}")


# =========================================================
# STREAMLIT UI
# =========================================================

st.title("📊 Census Data Extraction from Excel using LLM")
st.markdown("---")

# File uploader
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    # Save uploaded file to temporary location
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_file.write(uploaded_file.read())
        tmp_file_path = tmp_file.name
    
    try:
        if st.button("🚀 Extract Data", type="primary", use_container_width=True):
            progress_bar = st.progress(0)
            status_text = st.empty()
            log_messages = []
            
            def log_function(msg):
                status_text.text(msg)
                log_messages.append(msg)
                # Update progress if we can parse chunk info
                if "chunk" in msg.lower() and "/" in msg:
                    try:
                        import re
                        match = re.search(r'chunk (\d+)/(\d+)', msg.lower())
                        if match:
                            current, total = int(match.group(1)), int(match.group(2))
                            progress_bar.progress(current / total)
                    except:
                        pass
            
            # Process the file
            with st.spinner("Processing Excel file..."):
                results = process_combined_excel(tmp_file_path, log=log_function)
            
            progress_bar.progress(1.0)
            status_text.text("✅ Extraction completed!")
            
            if results:
                # Display extraction logs
                with st.expander("📋 Extraction Logs", expanded=False):
                    for msg in log_messages:
                        st.text(msg)
                
                # Convert to DataFrame for display
                df = pd.DataFrame(results)
                
                st.subheader("📊 Extracted Data")
                st.dataframe(df, width='stretch', height=400)
                
                # Statistics
                st.subheader("📈 Statistics")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Records", len(results))
                with col2:
                    # Count employees (relationship_to_employee = "Employee" or null/empty)
                    if 'relationship_to_employee' in df.columns:
                        employees = df['relationship_to_employee'].apply(
                            lambda x: str(x).lower().strip() in ['employee', 'ee', 'emp', 'e', ''] or pd.isna(x)
                        ).sum()
                        st.metric("Employees", int(employees))
                    else:
                        st.metric("Employees", "N/A")
                with col3:
                    if 'relationship_to_employee' in df.columns:
                        dependents = len(df) - employees
                        st.metric("Dependents", int(dependents))
                    else:
                        st.metric("Dependents", "N/A")
                
                # Download options
                st.subheader("💾 Download Results")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                base_name = uploaded_file.name.rsplit('.', 1)[0] if '.' in uploaded_file.name else uploaded_file.name
                
                # JSON download
                json_str = json.dumps(results, indent=2, ensure_ascii=False)
                json_filename = f"{base_name}_employees_{timestamp}.json"
                
                # CSV download
                csv_str = df.to_csv(index=False)
                csv_filename = f"{base_name}_employees_{timestamp}.csv"
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        "⬇️ Download JSON",
                        json_str,
                        json_filename,
                        "application/json",
                        use_container_width=True
                    )
                with col2:
                    st.download_button(
                        "⬇️ Download CSV",
                        csv_str,
                        csv_filename,
                        "text/csv",
                        use_container_width=True
                    )
                
                st.success(f"✅ Extraction completed successfully! Extracted {len(results)} records.")
            else:
                st.error("❌ No data extracted. Please check the extraction logs above.")
                if log_messages:
                    with st.expander("📋 Extraction Logs", expanded=True):
                        for msg in log_messages:
                            st.text(msg)
    
    finally:
        # Clean up temporary file
        if os.path.exists(tmp_file_path):
            os.unlink(tmp_file_path)

else:
    st.info("👆 Please upload an Excel file to begin extraction.")

# Footer
st.markdown("---")
footer_html = """
<div style="display: flex; justify-content: space-between; padding: 10px; font-size: 12px; color: #666;">
    <div>ConceptVine : Proof of Concept</div>
    <div>Powered by Shepardtri</div>
</div>
"""
st.markdown(footer_html, unsafe_allow_html=True)

