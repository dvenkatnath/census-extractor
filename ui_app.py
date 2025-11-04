import streamlit as st
import pandas as pd
import json, io, time, re, difflib
from datetime import datetime
from groq import Groq

# ---- CONFIG ----
import os
from dotenv import load_dotenv
load_dotenv()
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
if not GROQ_API_KEY:
    raise ValueError("❌ GROQ_API_KEY not found in .env file")
LLM_MODEL = "llama-3.3-70b-versatile"
LEARNING_FILE = "learned_mappings.json"
client = Groq(api_key=GROQ_API_KEY)

st.set_page_config("Census Extractor", layout="wide")
if "log" not in st.session_state:
    st.session_state.log = []

def log(msg):
    msg = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
    st.session_state.log.append(msg)
    st.sidebar.text("\n".join(st.session_state.log[-15:]))

# ---- LEARNING ----
def load_learnings():
    try:
        return json.load(open(LEARNING_FILE))
    except:
        return {}

def save_learnings(data):
    with open(LEARNING_FILE, "w") as f:
        json.dump(data, f, indent=2)

learned = load_learnings()

# ---- HEADER CLEANUP ----
def normalize(text):
    return re.sub(r"[^a-z0-9]+", "", str(text).lower().strip())

def clean_headers(df):
    first = list(df.columns)
    if any(re.match(r"^Unnamed", str(x)) for x in first):
        try:
            df.columns = [
                f"{str(a)} {str(b)}".strip()
                for a, b in zip(df.columns.get_level_values(0), df.columns.get_level_values(1))
            ]
        except Exception:
            df.columns = [str(c).strip() for c in first]
    else:
        df.columns = [str(c).strip() for c in first]
    df = df.dropna(how="all")
    return df

# ---- LLM MAPPER ----
def llm_map_columns(headers):
    joined = ", ".join(headers)
    prompt = f"""
You are an expert in mapping inconsistent Excel headers to standard census fields.
Given the headers below, map them to:
First Name, Last Name, DOB, Gender, Relationship, Dependent(Y/N),
Medical Coverage, Medical Plan, Dental Coverage, Dental Plan, Vision Coverage, Vision Plan, COBRA.

Return valid JSON like:
{{
  "First Name": "Employee Name",
  "DOB": "Date of Birth",
  "Gender": "Sex",
  "Medical Coverage": "Med Cov"
}}

Headers:
{joined}
"""
    log("🔍 Calling Groq LLaMA 3.3 for mapping...")
    try:
        resp = client.chat.completions.create(
            model=LLM_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=800,
        )
        content = resp.choices[0].message.content
        content = re.sub(r"```json|```", "", content).strip()
        mapping = json.loads(content)
        log("✅ Model mapping parsed successfully.")
        log(f"LLM Mapping Output: {json.dumps(mapping, indent=2)}")
        return mapping
    except Exception as e:
        log(f"⚠️ LLM mapping failed or invalid JSON: {e}")
        return {}

# ---- FUZZY MATCHER ----
def fuzzy_match_column(target, available_cols):
    target_norm = normalize(target)
    candidates = {col: normalize(col) for col in available_cols}
    best = difflib.get_close_matches(target_norm, candidates.values(), n=1, cutoff=0.6)
    if best:
        for orig, norm in candidates.items():
            if norm == best[0]:
                return orig
    return None

# ---- MAIN EXTRACTION ----
def process_excel(upload):
    xls = pd.ExcelFile(upload)
    all_dfs = []
    total_sheets = len(xls.sheet_names)
    progress = st.progress(0)
    for i, sheet in enumerate(xls.sheet_names):
        try:
            df = pd.read_excel(xls, sheet)
            df = clean_headers(df)
            headers = [str(c) for c in df.columns]
            log(f"📄 Sheet '{sheet}' headers: {headers}")
            llm_mapping = llm_map_columns(headers)
            if not llm_mapping:
                log("⚠️ No mapping returned by model. Skipping sheet.")
                continue

            mapped_df = pd.DataFrame()
            for field, llm_col in llm_mapping.items():
                match = fuzzy_match_column(llm_col, df.columns)
                if match:
                    mapped_df[field] = df[match]
                    log(f"✔️ {field} ← {match}")
                else:
                    log(f"❌ No match found for '{field}' (model suggested '{llm_col}')")

            if not mapped_df.empty:
                mapped_df["Sheet"] = sheet
                all_dfs.append(mapped_df)
            else:
                log(f"⚠️ No valid columns extracted from sheet '{sheet}'.")
        except Exception as e:
            log(f"⚠️ Error in sheet '{sheet}': {e}")
        progress.progress((i + 1) / total_sheets)

    progress.empty()
    if not all_dfs:
        return None
    final = pd.concat(all_dfs, ignore_index=True)
    save_learnings(learned)
    return final

# ---- STREAMLIT UI ----
st.title("📘 Census Extractor – Groq LLaMA 3.3 (Fuzzy Matching Enhanced)")
st.sidebar.header("🧠 Logs & Progress")
uploaded = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded:
    if st.button("🚀 Process File"):
        st.session_state.log.clear()
        log("Starting extraction process...")
        start = time.time()
        result = process_excel(uploaded)
        if result is None or result.empty:
            st.error("❌ No valid extraction found. Check logs for details.")
        else:
            st.success(f"✅ Extraction complete in {time.time()-start:.1f}s, Rows: {len(result)}")
            st.dataframe(result.head(50))
            out = io.BytesIO()
            result.to_excel(out, index=False)
            st.download_button("⬇️ Download Extracted Excel",
                               out.getvalue(),
                               file_name=f"{uploaded.name.split('.')[0]}_cleaned.xlsx")
            log("Extraction complete and file ready for download.")
else:
    st.info("Upload an Excel file to begin extraction.")
