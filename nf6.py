import pandas as pd
import json
import math
import time
import re
import csv
from groq import Groq
import tkinter as tk
from tkinter import Tk, filedialog, messagebox, ttk
import os
import threading
import subprocess
import platform

# =========================================================
# CONFIGURATION
# =========================================================
import os
from dotenv import load_dotenv
load_dotenv()
api_key = os.getenv("GROQ_API_KEY")
if not api_key:
    raise ValueError("❌ GROQ_API_KEY not found in .env file")
client = Groq(api_key=api_key)
LLM_MODEL = "llama-3.3-70b-versatile"
MAX_INPUT_CHARS = 40000  # per prompt
MAX_OUTPUT_CHECK = 40000  # if too short, request continuation

# =========================================================
# HELPER FUNCTIONS
# =========================================================

# def clean_json_output(output_text):
#     """Extracts valid JSON array/object from model output (handles ```json ... ```)."""
#     # Remove code block fences if present
#     match = re.search(r'```(?:json)?\s*([\s\S]*?)```', output_text)
#     if match:
#         output_text = match.group(1).strip()
#     # Find first [ and last ]
#     json_part = re.search(r'(\[.*\])', output_text, re.DOTALL)
#     if json_part:
#         output_text = json_part.group(1).strip()
#     return output_text
# def clean_json_output(output_text):
#     """Extract valid JSON content from LLM output even if mixed with extra text."""
#     output_text = output_text.strip()

#     # Remove markdown fences
#     output_text = re.sub(r"^```(?:json)?|```$", "", output_text, flags=re.MULTILINE).strip()

#     # Try direct JSON parse first
#     try:
#         json.loads(output_text)
#         return output_text
#     except:
#         pass

#     # Extract first JSON-like structure between [ ... ] or { ... }
#     match = re.search(r'(\[.*\]|\{.*\})', output_text, re.DOTALL)
#     if match:
#         return match.group(1).strip()

#     return output_text
def clean_json_output(output_text):
    """Extract and fix valid JSON content from messy LLM output."""
    import re, json

    if not output_text:
        return "[]"

    text = output_text.strip()

    # Remove markdown code fences
    text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.MULTILINE).strip()

    # Try direct parse first
    try:
        json.loads(text)
        return text
    except:
        pass

    # Extract all JSON-like objects/arrays
    matches = re.findall(r'(\[.*?\]|\{.*?\})', text, re.DOTALL)
    if not matches:
        return "[]"

    # Try each found JSON section
    valid_parts = []
    for m in matches:
        json_str = m

        # Basic fixes
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

def get_full_llm_output(prompt, model=None, log=None):
    """Send prompt to Groq and request continuation if truncated."""
    if model is None:
        model = LLM_MODEL
    
    full_output = ""
    part = 1

    while True:
        if log:
            log(f"🧠 Sending LLM request part {part}...")

        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=12288,
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


def process_excel_in_chunks(file_path, model=None, log=None):
    """Reads Excel file, extracts employee data JSON cleanly."""
    if model is None:
        model = LLM_MODEL
    
    all_results = []

    try:
        xls = pd.ExcelFile(file_path)
        if log:
            log(f"📘 Loaded workbook '{file_path}' with {len(xls.sheet_names)} sheets")

        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name)
            if df.empty:
                continue

            # Quick skip check: if no likely employee fields, skip
            combined_text = " ".join(df.columns.astype(str)) + " " + " ".join(df.astype(str).fillna("").values.flatten())
            if not any(keyword.lower() in combined_text.lower() for keyword in ["name", "dob", "zip", "employee", "dependent"]):
                if log:
                    log(f"⚠️ Skipping sheet '{sheet_name}' — no employee-like data detected")
                continue

            sheet_text = df.to_string(index=False)
            total_len = len(sheet_text)
            num_chunks = math.ceil(total_len / MAX_INPUT_CHARS)
            sheet_results = []

            if log:
                log(f"📄 Sheet '{sheet_name}' -> {num_chunks} chunks")

            for i in range(num_chunks):
                chunk_text = sheet_text[i * MAX_INPUT_CHARS : (i + 1) * MAX_INPUT_CHARS]

                prompt = f"""
                You are an expert data extractor.
                From the following table text (chunk {i+1}/{num_chunks} of sheet '{sheet_name}'),
                extract all employee-related records and dependent details in sequence.
                Return only valid JSON list (no extra text, no explanations).


                Fields:
                ["last_name","first_name","employee_name","home_zip_code",
                 "dob","gender","medical_coverage","medical_coverage_level","vision_coverage",
                 "vision_coverage_level","dental_coverage","dental_coverage_level",
                 "cobra_participation","relationship_to_employee","dependent_of_employee_row"]

                
                CRITICAL NAME EXTRACTION RULES:
                
                - **NAME SPLITTING**: If a column contains full names (e.g., "John Smith", "Mary Jane Doe"), you MUST split them into first_name and last_name:
                  * first_name: First word(s) of the full name (e.g., "John", "Mary Jane")
                  * last_name: Last word of the full name (e.g., "Smith", "Doe")
                  * employee_name: Keep the full name as is (e.g., "John Smith", "Mary Jane Doe")
                
                - **STANDARD NAME COLUMNS**: If there are separate "First Name" and "Last Name" columns, use them directly.
                
                - **MIXED FORMATS**: Handle both formats:
                  * Separate columns: Use first_name from "First Name" column, last_name from "Last Name" column
                  * Full name column: Split "Employee Name" or "Name" into first_name and last_name
                  * Both available: Prefer separate columns but ensure employee_name contains the full name
                
                - If First Name is missing or empty after extraction, skip that row.
                
                OTHER RULES:
                
                - Read all worksheets, find employee info, and fetch data once found do not process other sheets.
                - Use the exact field names provided.
                - If COBRA Participant is missing, return "No".
                - Dependents: split 'DEPENDENT n NAME' into First Name and Last Name using the same splitting logic.
                - Relationship to Employee: use exactly as shown in Excel (no grouping by last name).
                - Employee rows: Relationship to Employee = "Employee", Dependent of Employee Row = null.
                - Dependents: Dependent of Employee Row should link to the Employee in order of appearance (Slno).
                - Dependent data in separate row, not as sub node in json.
                - For medical coverage, dental coverage and vision coverage match and pick the plan name not provider.
                - Avoid processing sheet named "company info" and "enrollment info".
                - Read all the sheets data keep it in memory, find the employee data and dependent data and group it. Provide the output in meaningful family sequence.
                - Avoid timestamp on Date of Birth column, only provide available data.
                - Compensation type is not coverage level. Do not map.
                - Do not return NAN, return blank.
                - If dependent fields in another sheet, map it with the employee data based the first name and last name.
                - Do not map worker compensation code and salary to coverage and coverage level.

                Table:
                {chunk_text}
                """
                                # - include Sl.No in output json  - row starts with 1 - shold be a running number

                output_text = get_full_llm_output(prompt, model=model, log=log)

                cleaned = clean_json_output(output_text)

                try:
                    json_data = json.loads(cleaned)
                    if isinstance(json_data, list):
                        sheet_results.extend(json_data)
                    else:
                        sheet_results.append(json_data)
                except json.JSONDecodeError:
                    if log:
                        log(f"⚠️ Could not parse JSON in chunk {i+1}/{num_chunks}")
                    continue

                if log:
                    log(f"✅ Finished chunk {i+1}/{num_chunks} for sheet '{sheet_name}'")

                time.sleep(1)

            if sheet_results:
                all_results.extend(sheet_results)
                if log:
                    log(f"🧩 Combined {len(sheet_results)} records from '{sheet_name}'")

        if log:
            log(f"🎯 Total extracted records: {len(all_results)}")

        return all_results

    except Exception as e:
        if log:
            log(f"❌ Error: {e}")
        return []


def json_to_csv_by_input_name(json_filepath):
    """Converts JSON file (list of dicts) to CSV."""
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
# TKINTER UI APPLICATION
# =========================================================

class GroqCensusExtractorUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Census Data Extractor")
        self.root.geometry("1000x700")
        
        # Variables
        self.selected_file = tk.StringVar()
        self.selected_model = tk.StringVar(value="llama-3.3-70b-versatile")
        self.output_file_path = None
        
        # Create UI elements
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="LLM based Census data Extractor from Excel sheets", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection
        ttk.Label(main_frame, text="Select Excel File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.selected_file, width=50, state="readonly").grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5)
        
        # Button frame for Browse and Clear buttons
        file_button_frame = ttk.Frame(main_frame)
        file_button_frame.grid(row=1, column=2, pady=5)
        
        ttk.Button(file_button_frame, text="Browse", command=self.browse_file).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(file_button_frame, text="Clear", command=self.clear_messages).pack(side=tk.LEFT)
        
        # Model selection
        ttk.Label(main_frame, text="Select Model:").grid(row=2, column=0, sticky=tk.W, pady=5)
        model_frame = ttk.Frame(main_frame)
        model_frame.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5)
        model_frame.columnconfigure(0, weight=1)
        
        model_combo = ttk.Combobox(model_frame, textvariable=self.selected_model, state="readonly", width=47)
        model_combo['values'] = ("llama-3.3-70b-versatile", "llama-3.1-70b-versatile", "mixtral-8x7b-32768")
        model_combo.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Process button
        self.process_button = ttk.Button(main_frame, text="Process File", command=self.process_file, state="disabled")
        self.process_button.grid(row=3, column=0, columnspan=3, pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # Status text area
        ttk.Label(main_frame, text="Status:").grid(row=5, column=0, sticky=tk.W, pady=(10, 0))
        
        # Create frame for status text and scrollbar
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
        
        self.status_text = tk.Text(status_frame, height=15, wrap=tk.WORD)
        status_scrollbar = ttk.Scrollbar(status_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=status_scrollbar.set)
        
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        status_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # View output buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        
        self.view_json_button = ttk.Button(button_frame, text="View JSON Output", command=self.view_json, state="disabled")
        self.view_json_button.grid(row=0, column=0, padx=(0, 5), sticky=(tk.W, tk.E))
        
        self.view_csv_button = ttk.Button(button_frame, text="View CSV Output", command=self.view_csv, state="disabled")
        self.view_csv_button.grid(row=0, column=1, padx=(5, 0), sticky=(tk.W, tk.E))
        
        # Footer frame
        footer_frame = ttk.Frame(main_frame)
        footer_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(20, 0))
        footer_frame.columnconfigure(0, weight=1)
        
        # Get system theme and set appropriate color
        try:
            # Try to detect dark mode (macOS/Windows)
            import sys
            if sys.platform == "darwin":  # macOS
                # Check if system is in dark mode
                result = subprocess.run(['defaults', 'read', '-g', 'AppleInterfaceStyle'], 
                                     capture_output=True, text=True)
                is_dark = "Dark" in result.stdout
            else:
                # Default to light mode for other systems
                is_dark = False
        except:
            is_dark = False
        
        footer_color = "white" if is_dark else "black"
        
        # Left footer text - adaptive color
        left_footer = ttk.Label(footer_frame, text="ConceptVine : POC : Census Data Extraction", foreground=footer_color)
        left_footer.grid(row=0, column=0, sticky=tk.W)
        
        # Right footer text - adaptive color  
        right_footer = ttk.Label(footer_frame, text="Powered by Shepardtri", foreground=footer_color)
        right_footer.grid(row=0, column=1, sticky=tk.E)
        
        # Configure main frame grid weights
        main_frame.rowconfigure(6, weight=1)
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            # Clear previous messages
            self.status_text.delete(1.0, tk.END)
            self.selected_file.set(file_path)
            self.process_button.config(state="normal")
            self.log_status(f"File selected: {os.path.basename(file_path)}")
    
    def clear_messages(self):
        """Clear all status messages"""
        self.status_text.delete(1.0, tk.END)
        self.log_status("Messages cleared.")
            
    def log_status(self, message):
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.root.update()
        
    def process_file(self):
        if not self.selected_file.get():
            messagebox.showerror("Error", "Please select a file first.")
            return
            
        # Disable process button and start progress
        self.process_button.config(state="disabled")
        self.progress.start()
        
        # Start processing in a separate thread
        thread = threading.Thread(target=self.run_processing)
        thread.daemon = True
        thread.start()
        
    def run_processing(self):
        try:
            file_path = self.selected_file.get()
            model = self.selected_model.get()
            
            self.log_status(f"🚀 Starting extraction using model: {model}")
            self.log_status(f"✅ Model '{model}' connected successfully")
            self.log_status(f"⚙️ Config: MAX_INPUT_CHARS={MAX_INPUT_CHARS}, MAX_OUTPUT_CHECK={MAX_OUTPUT_CHECK}")
            self.log_status(f"📄 Processing file: {os.path.basename(file_path)}")
            
            # Set up output paths
            base, _ = os.path.splitext(file_path)
            self.output_json_path = base + "_employees.json"
            self.output_csv_path = base + "_employees.csv"
            
            # Process the file using the existing logic
            results = process_excel_in_chunks(file_path, model=model, log=self.log_status)
            
            if results:
                try:
                    with open(self.output_json_path, "w", encoding="utf-8") as f:
                        json.dump(results, f, indent=2, ensure_ascii=False)

                    json_to_csv_by_input_name(self.output_json_path)
                    self.log_status(f"\n✅ Final JSON saved to: {self.output_json_path}")
                    self.log_status(f"✅ Final CSV saved to: {self.output_csv_path}")
                    
                    # Show preview
                    df_out = pd.DataFrame(results).fillna("")
                    self.log_status("\n--- Extracted Data Preview ---")
                    self.log_status(df_out.head().to_string(index=False))
                    self.log_status("-------------------------------------")

                except Exception as e:
                    self.log_status(f"❌ Error saving JSON: {e}")
            else:
                self.log_status("⚠️ No data extracted.")
            
            # Enable view output buttons
            self.view_json_button.config(state="normal")
            self.view_csv_button.config(state="normal")
            
        except Exception as e:
            self.log_status(f"❌ Error: {e}")
            import traceback
            self.log_status(traceback.format_exc())
        finally:
            # Re-enable process button and stop progress
            self.progress.stop()
            self.process_button.config(state="normal")
            
    def view_json(self):
        if hasattr(self, 'output_json_path') and self.output_json_path and os.path.exists(self.output_json_path):
            try:
                if platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", self.output_json_path])
                elif platform.system() == "Windows":
                    os.startfile(self.output_json_path)
                else:  # Linux
                    subprocess.run(["xdg-open", self.output_json_path])
            except Exception as e:
                messagebox.showerror("Error", f"Could not open JSON file: {e}")
        else:
            messagebox.showerror("Error", "JSON output file not found.")
    
    def view_csv(self):
        if hasattr(self, 'output_csv_path') and self.output_csv_path and os.path.exists(self.output_csv_path):
            try:
                if platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", self.output_csv_path])
                elif platform.system() == "Windows":
                    os.startfile(self.output_csv_path)
                else:  # Linux
                    subprocess.run(["xdg-open", self.output_csv_path])
            except Exception as e:
                messagebox.showerror("Error", f"Could not open CSV file: {e}")
        else:
            messagebox.showerror("Error", "CSV output file not found.")

# =========================================================
# MAIN EXECUTION
# =========================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = GroqCensusExtractorUI(root)
    root.mainloop()
