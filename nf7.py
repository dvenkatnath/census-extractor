import pandas as pd
import json
import math
import time
import re
import csv
from groq import Groq
from tkinter import Tk, filedialog
import os
from dotenv import load_dotenv

# =========================================================
# CONFIGURATION
# =========================================================
load_dotenv()
api_key = os.getenv("GROQ_API_KEY")
if not api_key:
    raise ValueError("❌ GROQ_API_KEY not found in .env file")
client = Groq(api_key=api_key)
LLM_MODEL = "llama-3.3-70b-versatile"
MAX_INPUT_CHARS = 40000
MAX_OUTPUT_CHECK = 40000


# =========================================================
# HELPER FUNCTIONS
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
                - Read all worksheets, find employee info, and fetch data once found do not process other sheets .
                - Read all the rows in the input file ,in the excel work sheet if empty rows found skip and move to next row and read till the end of row with data.
                - Use the exact field names provided.
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
                - if relationship_to_employee not found leave blank
                - If employee data repeated in another sheet , do not reapet in output , employee row is unique also provide the dependent name in name column not the employee
                - provide employee name in respective dependent_of_employee_row
                - in coverage level , do not fill data splited from coverage. if coverage level not available provide blank

           

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
            time.sleep(1)

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
# MAIN EXECUTION
# =========================================================
if __name__ == "__main__":
    Tk().withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Excel file", filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not file_path:
        print("No file selected. Exiting.")
        exit()

    def console_log(msg): print(msg)
    results = process_combined_excel(file_path, log=console_log)

    output_file = file_path.replace(".xlsx", "_employees.json").replace(".xls", "_employees.json")
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    json_to_csv_by_input_name(output_file)
    print(f"\n✅ Extraction completed.\nJSON file saved to: {output_file}")