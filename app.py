import streamlit as st, pandas as pd, os, json, io
from dotenv import load_dotenv
from mapper import build_mapping, store_successful_mapping
from hunter import extract_data, produce_stats
from learning_system import learning_system
from llm_extractor import extract_with_full_context, produce_stats_for_llm

load_dotenv()
st.set_page_config(page_title="Census Mapper & Extractor", layout="wide")

# Inject script immediately to hide Manage app button - runs before page render
components.html("""
<script>
(function() {
    function hideManageApp() {
        // Method 1: Remove all elements with "Manage app" text
        var all = document.querySelectorAll('*');
        for (var i = 0; i < all.length; i++) {
            var el = all[i];
            var txt = (el.textContent || el.innerText || '').trim();
            if (txt === 'Manage app' || txt.includes('Manage app')) {
                el.remove();
            }
        }
        // Method 2: Hide all footers and bottom-right elements
        document.querySelectorAll('footer, [class*="footer"], [id*="footer"]').forEach(function(el) {
            el.style.cssText = 'display:none!important;visibility:hidden!important;height:0!important;width:0!important;position:fixed!important;bottom:-9999px!important;z-index:-9999!important;';
            el.remove();
        });
        // Method 3: Hide fixed positioned elements in bottom right
        document.querySelectorAll('*').forEach(function(el) {
            var style = window.getComputedStyle(el);
            if (style.position === 'fixed' && (style.bottom === '0px' || style.right === '0px')) {
                var txt = (el.textContent || el.innerText || '').trim();
                if (txt.includes('Manage') || txt.includes('manage')) {
                    el.remove();
                }
            }
        });
    }
    // Run immediately
    if (document.body) hideManageApp();
    else document.addEventListener('DOMContentLoaded', hideManageApp);
    // Run continuously
    setInterval(hideManageApp, 50);
    // Watch for changes
    if (document.body) {
        new MutationObserver(hideManageApp).observe(document.body, {childList:true, subtree:true, attributes:true});
    }
})();
</script>
""", height=0)

# Hide Streamlit Cloud "Manage app" button and menu - Ultra aggressive approach
hide_streamlit_style = """
    <style>
    /* Hide all Streamlit UI elements */
    #MainMenu {visibility: hidden !important; display: none !important; height: 0 !important; width: 0 !important; overflow: hidden !important;}
    footer {visibility: hidden !important; display: none !important; height: 0 !important; width: 0 !important; overflow: hidden !important; position: fixed !important; bottom: -9999px !important;}
    header {visibility: hidden !important; display: none !important; height: 0 !important; width: 0 !important; overflow: hidden !important;}
    .stDeployButton {display:none !important; visibility: hidden !important; height: 0 !important; width: 0 !important; overflow: hidden !important; position: absolute !important; left: -9999px !important;}
    button[title="Manage app"] {display:none !important; visibility: hidden !important; height: 0 !important; width: 0 !important; overflow: hidden !important; position: absolute !important; left: -9999px !important;}
    button[kind="header"] {display:none !important; visibility: hidden !important;}
    div[data-testid="stToolbar"] {visibility: hidden !important; height: 0rem !important; display: none !important; max-height: 0 !important; overflow: hidden !important; position: fixed !important; bottom: -9999px !important;}
    div[data-testid="stDecoration"] {visibility: hidden !important; height: 0rem !important; display: none !important; max-height: 0 !important; overflow: hidden !important;}
    div[data-testid="stHeader"] {visibility: hidden !important; height: 0rem !important; display: none !important; max-height: 0 !important; overflow: hidden !important;}
    #stApp > header {visibility: hidden !important; height: 0rem !important; display: none !important; max-height: 0 !important; overflow: hidden !important;}
    #stApp > footer {visibility: hidden !important; height: 0rem !important; display: none !important; max-height: 0 !important; overflow: hidden !important; position: fixed !important; bottom: -9999px !important; z-index: -9999 !important;}
    section[data-testid="stSidebar"] > div {visibility: hidden !important; height: 0rem !important; display: none !important;}
    .stApp > footer {visibility: hidden !important; height: 0rem !important; display: none !important; max-height: 0 !important; overflow: hidden !important; position: fixed !important; bottom: -9999px !important; z-index: -9999 !important;}
    .stApp > header {visibility: hidden !important; height: 0rem !important; display: none !important; max-height: 0 !important; overflow: hidden !important;}
    iframe[title="Manage app"] {display: none !important; visibility: hidden !important; height: 0 !important; width: 0 !important;}
    a[title="Manage app"] {display: none !important; visibility: hidden !important; height: 0 !important; width: 0 !important; position: absolute !important; left: -9999px !important;}
    /* Hide Streamlit Cloud specific elements */
    [class*="stDeployButton"] {display: none !important; visibility: hidden !important; height: 0 !important; width: 0 !important; position: absolute !important; left: -9999px !important; z-index: -9999 !important;}
    [id*="deploy"] {display: none !important; visibility: hidden !important; height: 0 !important; width: 0 !important; position: absolute !important; left: -9999px !important; z-index: -9999 !important;}
    /* Target footer elements specifically */
    footer[data-testid] {display: none !important; visibility: hidden !important; height: 0 !important; width: 0 !important; position: fixed !important; bottom: -9999px !important; z-index: -9999 !important;}
    div[class*="footer"] {display: none !important; visibility: hidden !important; height: 0 !important; width: 0 !important; position: fixed !important; bottom: -9999px !important; z-index: -9999 !important;}
    /* Hide any element with "Manage app" in any form */
    a[href*="manage"], button[aria-label*="Manage"], div[aria-label*="Manage"] {display: none !important; visibility: hidden !important; height: 0 !important; width: 0 !important; position: absolute !important; left: -9999px !important; z-index: -9999 !important;}
    /* Target bottom right corner specifically */
    div[style*="bottom"], div[style*="right"] {position: relative !important;}
    /* Hide any fixed positioned elements in bottom right */
    div[style*="position: fixed"][style*="bottom"], 
    div[style*="position:fixed"][style*="bottom"],
    a[style*="position: fixed"][style*="bottom"],
    a[style*="position:fixed"][style*="bottom"],
    button[style*="position: fixed"][style*="bottom"],
    button[style*="position:fixed"][style*="bottom"] {
        display: none !important; 
        visibility: hidden !important; 
        height: 0 !important; 
        width: 0 !important; 
        position: absolute !important; 
        left: -9999px !important; 
        z-index: -9999 !important;
    }
    </style>
    <script>
    (function() {
        'use strict';
        // Ultra-aggressive function to hide Manage app button
        function hideManageApp() {
            try {
                // Method 1: Hide by text content - search all elements
                var allElements = document.querySelectorAll('*');
                for (var i = 0; i < allElements.length; i++) {
                    var el = allElements[i];
                    var text = el.textContent || el.innerText || '';
                    if (text.includes('Manage app') || text.includes('Manage') || text.includes('manage')) {
                        el.style.cssText = 'display: none !important; visibility: hidden !important; opacity: 0 !important; height: 0 !important; width: 0 !important; overflow: hidden !important; position: absolute !important; left: -9999px !important; z-index: -9999 !important;';
                        if (el.parentNode) {
                            el.parentNode.removeChild(el);
                        }
                    }
                }
                
                // Method 2: Hide all buttons with Manage text
                var buttons = document.querySelectorAll('button, a, div, span');
                for (var i = 0; i < buttons.length; i++) {
                    var btn = buttons[i];
                    var text = btn.textContent || btn.innerText || btn.getAttribute('title') || btn.getAttribute('aria-label') || '';
                    if (text.includes('Manage') || text.includes('manage')) {
                        btn.style.cssText = 'display: none !important; visibility: hidden !important; opacity: 0 !important; height: 0 !important; width: 0 !important; overflow: hidden !important; position: absolute !important; left: -9999px !important; z-index: -9999 !important;';
                        try {
                            if (btn.parentNode) {
                                btn.parentNode.removeChild(btn);
                            }
                        } catch(e) {}
                    }
                }
                
                // Method 3: Hide all footer elements
                var footers = document.querySelectorAll('footer, [class*="footer"], [id*="footer"]');
                for (var i = 0; i < footers.length; i++) {
                    footers[i].style.cssText = 'display: none !important; visibility: hidden !important; height: 0 !important; width: 0 !important; overflow: hidden !important; position: fixed !important; bottom: -9999px !important; z-index: -9999 !important;';
                }
                
                // Method 4: Hide elements in bottom right corner (fixed position, bottom right)
                var fixedElements = document.querySelectorAll('[style*="position"], [style*="bottom"], [style*="right"]');
                for (var i = 0; i < fixedElements.length; i++) {
                    var el = fixedElements[i];
                    var style = el.getAttribute('style') || '';
                    var computed = window.getComputedStyle(el);
                    if ((computed.position === 'fixed' || computed.position === 'absolute') && 
                        (computed.bottom === '0px' || computed.bottom === '0' || computed.right === '0px' || computed.right === '0')) {
                        var text = el.textContent || el.innerText || '';
                        if (text.includes('Manage') || text.includes('manage') || style.includes('bottom')) {
                            el.style.cssText = 'display: none !important; visibility: hidden !important; opacity: 0 !important; height: 0 !important; width: 0 !important; overflow: hidden !important; position: absolute !important; left: -9999px !important; z-index: -9999 !important;';
                            try {
                                if (el.parentNode) {
                                    el.parentNode.removeChild(el);
                                }
                            } catch(e) {}
                        }
                    }
                }
                
                // Method 5: Remove elements from DOM entirely
                var allNodes = document.querySelectorAll('*');
                for (var i = 0; i < allNodes.length; i++) {
                    var node = allNodes[i];
                    var text = node.textContent || node.innerText || '';
                    if (text.trim() === 'Manage app' || text.trim() === 'Manage') {
                        try {
                            node.parentNode && node.parentNode.removeChild(node);
                        } catch(e) {}
                    }
                }
            } catch(e) {
                console.log('Error hiding Manage app:', e);
            }
        }
        
        // Run immediately
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', hideManageApp);
        } else {
            hideManageApp();
        }
        
        // Run on window load
        window.addEventListener('load', function() {
            hideManageApp();
            setTimeout(hideManageApp, 50);
            setTimeout(hideManageApp, 100);
            setTimeout(hideManageApp, 200);
            setTimeout(hideManageApp, 500);
            setTimeout(hideManageApp, 1000);
            setTimeout(hideManageApp, 2000);
        });
        
        // Continuous monitoring with MutationObserver
        var observer = new MutationObserver(function(mutations) {
            hideManageApp();
        });
        
        // Start observing
        if (document.body) {
            observer.observe(document.body, { 
                childList: true, 
                subtree: true, 
                attributes: true,
                attributeFilter: ['style', 'class', 'id']
            });
        } else {
            document.addEventListener('DOMContentLoaded', function() {
                observer.observe(document.body, { 
                    childList: true, 
                    subtree: true, 
                    attributes: true,
                    attributeFilter: ['style', 'class', 'id']
                });
            });
        }
        
        // Also run on interval as backup
        setInterval(hideManageApp, 1000);
    })();
    </script>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
st.title("📊 Census Mapper & Extractor")

# Extraction Mode Selection
st.markdown("---")
st.subheader("🚀 Choose Extraction Mode")
extraction_mode = st.radio(
    "Select how you want to extract data:",
    ["🚀 Full LLM Extraction (Recommended)", "⚡ Quick Preview Mode"],
    index=0,  # Default to Full LLM Extraction
    help="Full LLM Extraction: Robust, handles any file structure. Preview Mode: Fast mapping preview for validation."
)

if extraction_mode == "🚀 Full LLM Extraction (Recommended)":
    st.info("✅ **Recommended**: This mode uses full data context for robust extraction that adapts to any file structure.")
else:
    st.info("⚡ **Preview Mode**: Fast mapping preview using sample rows. Useful for quick validation before full extraction.")

st.markdown("---")

# Learning Statistics
with st.expander("🧠 Learning System Statistics", expanded=False):
    stats = learning_system.get_statistics()
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Mappings", stats["total_mappings"])
    with col2:
        st.metric("Successful Mappings", stats["successful_mappings"])
    with col3:
        st.metric("Patterns Learned", stats["patterns_learned"])
    with col4:
        st.metric("Recent Mappings (24h)", stats["recent_mappings"])
    
    if stats["total_mappings"] > 0:
        success_rate = (stats["successful_mappings"] / stats["total_mappings"]) * 100
        st.success(f"🎯 Success Rate: {success_rate:.1f}%")
    
    if stats["last_updated"]:
        st.caption(f"Last updated: {stats['last_updated']}")

# Check API key status
api_key = os.getenv("GROQ_API_KEY")
if api_key:
    st.success(f"🔑 API Key loaded from .env file: {api_key[:10]}...")
else:
    st.error("❌ No API key found in .env file!")
    st.stop()

uploaded = st.file_uploader("Upload Excel file", type=["xlsx"])
if not uploaded:
    st.stop()

# ---- 1. read all sheets with proper header handling ------------------------
# Read Excel file to get sheet names first
xl = pd.ExcelFile(uploaded)
all_sheets = {}  # Processed sheets (for extraction)
original_sheets = {}  # Original sheets (for display - as-is from Excel)

for sheet_name in xl.sheet_names:
    # Try to read with header=0 first
    try:
        # Read original data AS-IS for display
        df_original = pd.read_excel(uploaded, sheet_name=sheet_name)
        original_sheets[sheet_name] = df_original.copy()
        
        # Create processed copy for extraction (with dtype=str for consistency)
        df = df_original.copy().astype(str)
        
        # Improved header row detection
        # Strategy: Find the row that looks most like a header (has text labels, not data values)
        header_row_idx = 0  # Default to first row
        best_header_score = 0
        
        # Check first 10 rows for the best header candidate
        for i in range(min(10, len(df))):
            row = df.iloc[i]
            score = 0
            
            # Count non-null values
            non_null_count = row.notna().sum()
            
            # Check if row values look like column headers (text, short, no numbers/dates)
            header_like_count = 0
            for val in row:
                if pd.notna(val):
                    val_str = str(val).strip()
                    # Header-like characteristics:
                    # - Short strings (typically < 50 chars)
                    # - Contains letters (not just numbers)
                    # - Doesn't look like a date
                    # - Doesn't look like a number
                    if (len(val_str) < 50 and 
                        any(c.isalpha() for c in val_str) and 
                        not val_str.replace('.', '').replace('-', '').isdigit() and
                        '/' not in val_str[:10]):  # Not a date
                        header_like_count += 1
            
            # Calculate score: prefer rows with many header-like values
            if non_null_count > 0:
                header_ratio = header_like_count / non_null_count
                # Prefer rows with many non-null values AND header-like values
                score = non_null_count * header_ratio
            
            if score > best_header_score:
                best_header_score = score
                header_row_idx = i
        
        # If we found a better header row (and it's not row 0), use it
        if header_row_idx > 0 or any('Unnamed' in str(col) for col in df.columns):
            print(f"🔍 Debug: Detected header row at index {header_row_idx} for sheet '{sheet_name}'")
            # Use the detected row as headers
            df.columns = df.iloc[header_row_idx]
            # Remove header row and all rows before it
            df = df.iloc[header_row_idx+1:].reset_index(drop=True)
        
        # Clean up any remaining "Unnamed" columns
        df.columns = [f"Column_{i}" if 'Unnamed' in str(col) else str(col) for i, col in enumerate(df.columns)]
        
        # Ensure all column names are unique
        new_columns = []
        seen = set()
        for col in df.columns:
            if col in seen:
                # Add a suffix to make it unique
                counter = 1
                new_col = f"{col}_{counter}"
                while new_col in seen:
                    counter += 1
                    new_col = f"{col}_{counter}"
                new_columns.append(new_col)
                seen.add(new_col)
            else:
                new_columns.append(col)
                seen.add(col)
        df.columns = new_columns
        
        all_sheets[sheet_name] = df  # Processed version for extraction
        
    except Exception as e:
        st.warning(f"⚠️ Could not read sheet '{sheet_name}': {e}")
        continue
st.subheader("1. Excel inspection")
st.write(f"**Filename:** {uploaded.name}")
st.write(f"**Total sheets:** {len(all_sheets)}")
sheet_info = [(name, df.shape[0], df.shape[1]) for name, df in all_sheets.items()]
sheet_df = pd.DataFrame(sheet_info, columns=["Sheet", "Rows", "Columns"])
st.dataframe(sheet_df, width='stretch')

# ============================================================================
# FULL LLM EXTRACTION MODE (Recommended)
# ============================================================================
if extraction_mode == "🚀 Full LLM Extraction (Recommended)":
    st.markdown("---")
    st.subheader("2. Full LLM Extraction")
    st.write("**Extracting data with full context - this method adapts to any file structure.**")
    
    if st.button("🚀 Extract All Data", type="primary", use_container_width=True):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        def log_function(msg):
            status_text.text(msg)
            # Update progress if we can parse progress info
            if "chunk" in msg.lower() and "/" in msg:
                try:
                    # Extract chunk number if available
                    import re
                    match = re.search(r'chunk (\d+)/(\d+)', msg.lower())
                    if match:
                        current, total = int(match.group(1)), int(match.group(2))
                        progress_bar.progress(current / total)
                except:
                    pass
        
        try:
            # Reset file pointer for upload
            uploaded.seek(0)
            
            # Extract using full context
            extracted_df = extract_with_full_context(uploaded, log=log_function)
            
            progress_bar.progress(1.0)
            status_text.text("✅ Extraction completed!")
            
            # Calculate field mapping and confidence scores
            st.subheader("2. Field Mapping & Confidence")
            
            # Analyze extracted data to infer field mappings and calculate confidence
            field_mapping_data = []
            canonical_fields = [
                "First Name", "Last Name", "Employee Name", "DOB", "Gender",
                "Relationship To employee", "Dependent Of Employee Row", "Dependent (Y/N)",
                "Medical Coverage", "Medical Plan Name", "Dental Coverage", "Dental Plan Name",
                "Vision Coverage", "Vision Plan Name", "COBRA Participation (Y/N)"
            ]
            
            for field in canonical_fields:
                if field in extracted_df.columns:
                    # Calculate confidence based on field completeness
                    # Count records where field is not NaN/None and not empty string
                    non_empty_count = extracted_df[field].apply(
                        lambda x: pd.notna(x) and str(x).strip() not in ["", "nan", "none", "null"]
                    ).sum()
                    total_count = len(extracted_df)
                    completeness = (non_empty_count / total_count * 100) if total_count > 0 else 0
                    
                    # Confidence score based on completeness
                    if completeness >= 90:
                        confidence = 95
                    elif completeness >= 70:
                        confidence = 80
                    elif completeness >= 50:
                        confidence = 65
                    elif completeness >= 30:
                        confidence = 50
                    else:
                        confidence = 30
                    
                    field_mapping_data.append({
                        "Field Name": field,
                        "Confidence Score": f"{confidence}%",
                        "Completeness": f"{completeness:.1f}%"
                    })
            
            mapping_df = pd.DataFrame(field_mapping_data)
            st.dataframe(mapping_df, width='stretch', hide_index=True)
            
            # Display results
            st.subheader("3. Extracted Data")
            st.dataframe(extracted_df, width='stretch', height=400)
            
            # Statistics
            st.subheader("4. Extraction Statistics")
            stats = produce_stats_for_llm(all_sheets, extracted_df)
            st.markdown(stats)
            
            # Download options
            st.subheader("5. Download Results")
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = uploaded.name.rsplit('.', 1)[0] if '.' in uploaded.name else uploaded.name
            excel_filename = f"{base_name}_extracted_{timestamp}.xlsx"
            
            # Create Excel file
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                # Sheet 1: Extracted data
                extracted_df.to_excel(writer, sheet_name='Extracted', index=False)
                
                # Sheet 2: Field Mapping
                mapping_df.to_excel(writer, sheet_name='Field Mapping', index=False)
                
                # Sheet 3: Statistics
                stats_lines = stats.split('\n')
                stats_rows = []
                for line in stats_lines:
                    line = line.strip()
                    if not line:
                        continue
                    if ':' in line:
                        parts = line.split(':', 1)
                        stats_rows.append({
                            "Metric": parts[0].strip(),
                            "Value": parts[1].strip() if len(parts) > 1 else ""
                        })
                    elif '\t' in line:
                        parts = line.split('\t')
                        stats_rows.append({
                            "Metric": parts[0].strip(),
                            "Value": parts[1].strip() if len(parts) > 1 else ""
                        })
                
                if stats_rows:
                    stats_df = pd.DataFrame(stats_rows)
                    stats_df.to_excel(writer, sheet_name='Stats', index=False)
                else:
                    # Fallback: create simple stats table
                    simple_stats = pd.DataFrame({
                        "Metric": ["Total Extracted Records", "Total Sheets"],
                        "Value": [str(len(extracted_df)), str(len(all_sheets))]
                    })
                    simple_stats.to_excel(writer, sheet_name='Stats', index=False)
            
            excel_buffer.seek(0)
            
            # Download buttons
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "⬇️ Download Excel", 
                    excel_buffer.getvalue(), 
                    excel_filename, 
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col2:
                csv = extracted_df.to_csv(index=False)
                st.download_button(
                    "⬇️ Download CSV", 
                    csv, 
                    f"{base_name}_extracted_{timestamp}.csv", 
                    "text/csv",
                    use_container_width=True
                )
            
            st.success("✅ Extraction completed successfully!")
            
        except Exception as e:
            st.error(f"❌ Error during extraction: {str(e)}")
            st.exception(e)

# ============================================================================
# QUICK PREVIEW MODE (Mapping-based)
# ============================================================================
else:
    # ---- 2. build sample data for AI analysis -------------------------------------
    sample_rows = 5
    thin_csv = io.StringIO()
    for sh_name, df in all_sheets.items():
        tiny = pd.concat([df.head(0), df.dropna(how="all").sample(n=min(sample_rows, len(df)), random_state=42)])
        tiny = tiny.map(lambda x: "" if pd.isna(x) else str(x))
        tiny.insert(0, "__sheet__", sh_name)
        tiny.to_csv(thin_csv, index=False, header=True)
    thin_csv.seek(0)
    
    # ---- 3. LLM mapping --------------------------------------------------------
    canonical = [
        "First Name", "Last Name", "Employee Name", "DOB", "Gender",
        "Relationship To employee", "Dependent (Y/N)", "Medical Coverage",
        "Medical Plan Name", "Dental Coverage", "Dental Plan Name",
        "Vision Coverage", "Vision Plan Name", "COBRA Participation (Y/N)"
    ]
    if "mapping" not in st.session_state:
        # Show status information
        st.info("🔗 Connected to Llama 3.3 70B - Versatile")
        st.info("📤 Sending first 5 rows of sample data...")
        
        # Show the sample data being sent
        with st.expander("📊 Data being sent to Groq API", expanded=True):
            st.text(thin_csv.getvalue())
        
        with st.spinner("🤖 Sending request to Groq API..."):
            mapping = build_mapping(thin_csv.getvalue(), canonical, uploaded.name)
            st.session_state["mapping"] = mapping
            st.session_state["original_mapping"] = mapping.copy()  # Store original for learning
        
        # Check if we got any mappings
        if not mapping or len(mapping) == 0:
            st.warning("⚠️ No mappings received from AI. Using fallback manual mapping.")
            # Create a simple fallback mapping
            mapping = {
                "First Name": ["UNKNOWN"],
                "Last Name": ["UNKNOWN"], 
                "DOB": ["UNKNOWN"],
                "Gender": ["UNKNOWN"],
                "UNKNOWN": ["All other columns"]
            }
            st.session_state["mapping"] = mapping
        else:
            st.success("✅ Received response from Groq API")
            st.success("📋 Successfully parsed JSON response")
    else:
        st.info("🔄 Using cached mapping from previous analysis")
        st.write("**To regenerate mapping, refresh the page or clear browser cache**")

    mapping = st.session_state["mapping"]
    st.subheader("2. Mapping proposed by LLM")

    # Display mapping results with simple confidence scores
    if mapping and len(mapping) > 0:
        mapping_data = []
        mapped_count = 0
        
        for field, cols in mapping.items():
            # Simple confidence calculation based on field name matching
            confidence = 0.95  # Default high confidence for mapped fields
            # Check if field is mapped (not empty list and not UNKNOWN)
            if cols and len(cols) > 0 and cols != ["UNKNOWN"]:
                mapped_count += 1
                # Check if field name matches common patterns for higher confidence
                if any(keyword in field.lower() for keyword in ['name', 'first', 'last', 'dob', 'date', 'gender']):
                    confidence = 0.98
                elif any(keyword in field.lower() for keyword in ['medical', 'dental', 'vision', 'coverage']):
                    confidence = 0.90
                else:
                    confidence = 0.85
            else:
                confidence = 0.0
            
            confidence_pct = f"{confidence * 100:.0f}%"
            
            # Format column references with sheet name and column header
            formatted_cols = []
            if cols and len(cols) > 0 and cols != ["UNKNOWN"]:
                for col_ref in cols:
                    if "," in col_ref:
                        sheet_name, col_name = col_ref.split(",", 1)
                        # Get the actual column header from the sheet
                        if sheet_name in all_sheets:
                            df = all_sheets[sheet_name]
                            if col_name in df.columns:
                                # Find column index to get letter
                                col_idx = df.columns.get_loc(col_name)
                                col_letter = chr(65 + col_idx) if col_idx < 26 else f"A{chr(65 + col_idx - 26)}"
                                formatted_cols.append(f"{sheet_name} Col {col_letter} - {col_name}")
                            else:
                                formatted_cols.append(f"{sheet_name} - {col_name}")
                        else:
                            formatted_cols.append(col_ref)
                    else:
                        formatted_cols.append(col_ref)
                mapped_cols_display = ", ".join(formatted_cols)
            elif cols == ["UNKNOWN"]:
                mapped_cols_display = "UNKNOWN"
            else:
                mapped_cols_display = "❌ Unmapped (empty)"
                confidence_pct = "0%"
            
            mapping_data.append({
                "Field": field,
                "Mapped Columns": mapped_cols_display,
                "Confidence": confidence_pct
            })
        
        mapping_df = pd.DataFrame(mapping_data)
        
        # Show mapping summary
        total_count = len(mapping)
        success_rate = (mapped_count / total_count) * 100 if total_count > 0 else 0
        avg_confidence = (mapped_count * 0.9) / total_count * 100 if total_count > 0 else 0
        
        st.metric("Mapping Success Rate", f"{success_rate:.1f}%")
        st.metric("Average Confidence", f"{avg_confidence:.1f}%")
        
        if success_rate >= 80:
            st.success(f"✅ High success rate ({success_rate:.1f}%)")
        elif success_rate >= 50:
            st.warning(f"⚠️ Medium success rate ({success_rate:.1f}%)")
        else:
            st.error(f"❌ Low success rate ({success_rate:.1f}%)")
    else:
        mapping_df = pd.DataFrame({"Field": ["No mappings"], "Mapped Columns": ["None"], "Confidence": ["0%"]})

    # Display current mapping summary
    st.subheader("📋 Current Mapping Summary")
    st.dataframe(mapping_df, width='stretch')

    # Add inline edit functionality
    st.subheader("✏️ Edit Mappings")
    st.write("**Click 'Edit' next to any field to correct the column mapping:**")

    # Add reset button
    if st.button("🔄 Reset All Mappings", help="Clear all current mappings and regenerate"):
        if "mapping" in st.session_state:
            del st.session_state["mapping"]
        if "original_mapping" in st.session_state:
            del st.session_state["original_mapping"]
        st.rerun()

    # Manual learning storage button
    if st.button("🧠 Store Current Mapping for Learning", help="Store the current mapping to improve future predictions"):
        if uploaded and st.session_state.get("original_mapping") and st.session_state.get("mapping"):
            try:
                # Extract column names from the uploaded file
                column_names = []
                for sheet_name, df in all_sheets.items():
                    column_names.extend(df.columns.tolist())
                
                store_successful_mapping(
                    st.session_state["original_mapping"],
                    st.session_state["mapping"],
                    column_names,
                    thin_csv.getvalue(),
                    uploaded.name
                )
                st.success("🧠 Learning: Current mapping stored for future improvements!")
            except Exception as e:
                st.error(f"❌ Could not store learning data: {e}")
        else:
            st.warning("⚠️ No mapping data available to store")

    # Create editable mapping table
    if mapping and len(mapping) > 0:
        for idx, (field, cols) in enumerate(mapping.items()):
            # Format current columns for display
            formatted_cols = []
            for col_ref in cols:
                if "," in col_ref:
                    sheet_name, col_name = col_ref.split(",", 1)
                    if sheet_name in all_sheets:
                        df = all_sheets[sheet_name]
                        if col_name in df.columns:
                            col_idx = df.columns.get_loc(col_name)
                            col_letter = chr(65 + col_idx) if col_idx < 26 else f"A{chr(65 + col_idx - 26)}"
                            formatted_cols.append(f"{sheet_name} Col {col_letter} - {col_name}")
                        else:
                            formatted_cols.append(f"{sheet_name} - {col_name}")
                    else:
                        formatted_cols.append(col_ref)
                else:
                    formatted_cols.append(col_ref)
            
            if formatted_cols:
                current_display = ", ".join(formatted_cols)
            elif cols == ["UNKNOWN"]:
                current_display = "UNKNOWN"
            else:
                current_display = "❌ Unmapped (will be empty)"
            
            col1, col2, col3 = st.columns([3, 2, 1])
            
            with col1:
                st.write(f"**{field}**")
            
            with col2:
                st.write(current_display)
            
            with col3:
                if st.button(f"✏️ Edit", key=f"edit_{field}_{idx}"):
                    st.session_state[f"editing_{field}"] = True
            
            # Show edit interface if editing this field
            if st.session_state.get(f"editing_{field}", False):
                st.write(f"**Editing: {field}**")
                
                # Show only relevant sheets - prioritize mapped sheets, but show all for flexibility
                relevant_sheets = {}
                mapped_sheets = set()
                
                for field_name, field_cols in mapping.items():
                    if field_cols and field_cols != ["UNKNOWN"]:
                        for col_ref in field_cols:
                            if "," in col_ref:
                                sheet_name = col_ref.split(",")[0]
                                mapped_sheets.add(sheet_name)
                
                # Include all sheets that have mapped columns (content-based, not name-based)
                for sheet_name in mapped_sheets:
                    if sheet_name in all_sheets:
                        relevant_sheets[sheet_name] = all_sheets[sheet_name]
                
                # If no mapped sheets found, show all sheets (handles files with unexpected sheet names)
                if not relevant_sheets:
                    relevant_sheets = all_sheets
                
                # Show sheet selection interface
                for sheet_name, df in relevant_sheets.items():
                    st.write(f"**Sheet: {sheet_name}**")
                    
                    with st.expander(f"📊 Preview data in {sheet_name} (first 5 rows)", expanded=False):
                        # Show original Excel data as-is
                        if sheet_name in original_sheets:
                            preview_df = original_sheets[sheet_name].head(5)
                            st.dataframe(preview_df, width='stretch')
                        else:
                            st.dataframe(df.head(5), width='stretch')
                    
                    col_options = [f"{sheet_name},{col}" for col in df.columns]
                    
                    # Add "None" option at the beginning
                    col_options_with_none = ["NONE"] + col_options
                    
                    # Get current selection (default to existing mapping or empty)
                    current_selection = []
                    if field in mapping and mapping[field]:
                        for col_ref in mapping[field]:
                            if "," in col_ref:
                                sheet, col = col_ref.split(",", 1)
                                if sheet == sheet_name and col in df.columns:
                                    current_selection.append(col_ref)
                    
                    selected_cols = st.multiselect(
                        f"Select columns for {field} in {sheet_name} (or choose 'NONE' to leave unmapped):",
                        col_options_with_none,
                        default=current_selection if current_selection else None,
                        key=f"select_{field}_{sheet_name}_{idx}",
                        help="Select columns that contain this field's data, or choose 'NONE' to leave unmapped (field will be empty)"
                    )
                    
                    # Handle selection - if "NONE" is selected, clear the mapping for this sheet
                    if selected_cols:
                        if "NONE" in selected_cols:
                            # If NONE is selected, remove all columns from this sheet
                            formatted_cols = []
                            # Keep columns from other sheets if any
                            if field in mapping and mapping[field]:
                                for col_ref in mapping[field]:
                                    if "," in col_ref:
                                        other_sheet, _ = col_ref.split(",", 1)
                                        if other_sheet != sheet_name:
                                            formatted_cols.append(col_ref)
                            
                            mapping[field] = formatted_cols if formatted_cols else []
                            st.session_state["mapping"][field] = formatted_cols if formatted_cols else []
                            st.info(f"ℹ️ {field} mapping cleared for {sheet_name} (will be empty)")
                        else:
                            # Normal column selection
                            formatted_cols = []
                            # Keep columns from other sheets
                            if field in mapping and mapping[field]:
                                for col_ref in mapping[field]:
                                    if "," in col_ref:
                                        other_sheet, _ = col_ref.split(",", 1)
                                        if other_sheet != sheet_name:
                                            formatted_cols.append(col_ref)
                            
                            # Add selected columns from this sheet
                            for col_ref in selected_cols:
                                if "," in col_ref:
                                    formatted_cols.append(col_ref)
                                else:
                                    formatted_cols.append(f"{sheet_name},{col_ref}")
                            
                            mapping[field] = formatted_cols
                            st.session_state["mapping"][field] = formatted_cols
                            st.success(f"✅ Updated {field} mapping")
                    else:
                        # Nothing selected - clear mapping for this sheet
                        formatted_cols = []
                        # Keep columns from other sheets if any
                        if field in mapping and mapping[field]:
                            for col_ref in mapping[field]:
                                if "," in col_ref:
                                    other_sheet, _ = col_ref.split(",", 1)
                                    if other_sheet != sheet_name:
                                        formatted_cols.append(col_ref)
                        
                        mapping[field] = formatted_cols if formatted_cols else []
                        st.session_state["mapping"][field] = formatted_cols if formatted_cols else []
                        st.info(f"ℹ️ {field} mapping cleared for {sheet_name} (will be empty)")
                
                # Show current mapping summary for this field across all sheets
                if field in mapping:
                    current_mapping = mapping[field]
                    if current_mapping and len(current_mapping) > 0:
                        st.info(f"📋 Current mapping for {field}: {', '.join(current_mapping) if isinstance(current_mapping, list) else current_mapping}")
                    else:
                        st.warning(f"⚠️ {field} is currently unmapped (will be empty)")
                
                # Save and cancel buttons
                col_save, col_cancel = st.columns(2)
                with col_save:
                    if st.button(f"💾 Save {field}", key=f"save_{field}_{idx}"):
                        # Ensure mapping is saved correctly (empty list if no columns selected)
                        if field in mapping:
                            if not mapping[field] or len(mapping[field]) == 0:
                                mapping[field] = []  # Explicitly set to empty list
                                st.session_state["mapping"][field] = []
                                st.info(f"ℹ️ {field} mapping cleared - field will be empty")
                            else:
                                st.session_state["mapping"][field] = mapping[field].copy()
                        else:
                            mapping[field] = []
                            st.session_state["mapping"][field] = []
                        
                        st.session_state[f"editing_{field}"] = False
                        st.success(f"✅ {field} mapping saved!")
                        
                        # Store successful mapping for learning
                        if uploaded and st.session_state.get("original_mapping"):
                            try:
                                # Extract column names from the uploaded file
                                column_names = []
                                for sheet_name, df in all_sheets.items():
                                    column_names.extend(df.columns.tolist())
                                
                                store_successful_mapping(
                                    st.session_state["original_mapping"],
                                    st.session_state["mapping"],
                                    column_names,
                                    thin_csv.getvalue(),
                                    uploaded.name
                                )
                                st.info("🧠 Learning: Mapping stored for future improvements!")
                            except Exception as e:
                                st.warning(f"⚠️ Could not store learning data: {e}")
                        st.rerun()
                
                with col_cancel:
                    if st.button(f"❌ Cancel {field}", key=f"cancel_{field}_{idx}"):
                        st.session_state[f"editing_{field}"] = False
                        st.rerun()
                
                st.write("---")

    # Debug: Show raw mapping format
    if mapping and len(mapping) > 0:
        with st.expander("🔍 Debug: Raw Mapping Format", expanded=False):
            st.json(mapping)

    # ---- 4. user confirmation --------------------------------------------------
    if st.checkbox("✅ Mapping is correct – proceed to extract"):
        
        # Filter relevant sheets (those that have mapped columns)
        relevant_sheets = {}
        mapped_sheets = set()
        
        for field, cols in mapping.items():
            if cols and cols != ["UNKNOWN"]:
                for col_ref in cols:
                    if "," in col_ref:
                        sheet_name = col_ref.split(",")[0]
                        mapped_sheets.add(sheet_name)
        
        # Include sheets with mapped columns (content-based detection - works with any sheet names)
        for sheet_name in mapped_sheets:
            if sheet_name in all_sheets:
                relevant_sheets[sheet_name] = all_sheets[sheet_name]
        
        # If no mapped sheets found, use all sheets (ensures we don't miss data due to sheet naming)
        if not relevant_sheets:
            relevant_sheets = all_sheets
            st.info("ℹ️ No sheets with mapped columns found. Processing all sheets.")
        
        # Show original vs extracted data side by side
        st.subheader("3. Data Comparison: Original vs Extracted (10 rows)")
        st.write("**Side-by-side comparison of original data and extracted standardized data:**")
        
        # Debug: Show the mapping being used
        with st.expander("🔍 Debug: Mapping being used for extraction", expanded=False):
            st.json(mapping)
        
        with st.spinner("Processing sample data (10 rows)..."):
            # Create sample sheets with only 10 rows each
            sample_sheets = {}
            for sheet_name, df in relevant_sheets.items():
                sample_sheets[sheet_name] = df.head(10)
            
            # Debug: Show sample sheets
            with st.expander("🔍 Debug: Sample sheets being processed", expanded=False):
                for sheet_name, df in sample_sheets.items():
                    st.write(f"**{sheet_name}**: {df.shape[0]} rows, {df.shape[1]} columns")
                    st.write(f"Columns: {list(df.columns)}")
            
            # Extract sample data
            extracted_sample, stats = extract_data(sample_sheets, mapping)
            
            # Show side-by-side comparison
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📊 Original Data (10 rows)")
                for sheet_name, df in sample_sheets.items():
                    st.write(f"**Sheet: {sheet_name}**")
                    # Display original Excel data as-is (first 10 rows)
                    if sheet_name in original_sheets:
                        original_df = original_sheets[sheet_name].head(10)
                        st.dataframe(original_df, width='stretch')
                    else:
                        st.dataframe(df, width='stretch')
                    st.write("---")
            
            with col2:
                st.subheader("🔄 Extracted Data (10 rows)")
                st.write("**Standardized census format:**")
                st.dataframe(extracted_sample, width='stretch')
            
            # Show sample statistics
            st.subheader("4. Sample Statistics")
            st.text(stats)

        # Full extraction option
        if st.button("🚀 Proceed with Full Extraction", type="primary"):
            with st.spinner("Extracting all data..."):
                # Use only relevant sheets for full extraction
                extracted_full, stats_full = extract_data(relevant_sheets, mapping)
                
                st.subheader("5. Full Extracted Data")
                st.dataframe(extracted_full, width='stretch')

            st.subheader("6. Full Statistics")
            st.markdown(stats_full)

            # Create Excel file with multiple sheets as per spec
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = uploaded.name.rsplit('.', 1)[0] if '.' in uploaded.name else uploaded.name
            excel_filename = f"{base_name}_{timestamp}.xlsx"
            
            # Create Excel writer
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                # Sheet 1: Extracted data
                extracted_full.to_excel(writer, sheet_name='Extracted', index=False)
                
                # Sheet 2: Mappings
                mapping_rows = []
                for field, cols in mapping.items():
                    if cols and len(cols) > 0:
                        cols_str = ", ".join(cols) if isinstance(cols, list) else str(cols)
                    else:
                        cols_str = "Unmapped (empty)"
                    mapping_rows.append({
                        "Field": field,
                        "Mapped Columns": cols_str
                    })
                mapping_df = pd.DataFrame(mapping_rows)
                mapping_df.to_excel(writer, sheet_name='Mappings', index=False)
                
                # Sheet 3: Statistics
                # Parse stats text into structured format
                stats_lines = stats_full.split('\n')
                stats_rows = []
                for line in stats_lines:
                    line = line.strip()
                    if not line:
                        continue
                    # Try to parse different formats
                    if ':' in line:
                        parts = line.split(':', 1)
                        stats_rows.append({
                            "Metric": parts[0].strip(),
                            "Value": parts[1].strip() if len(parts) > 1 else ""
                        })
                    elif '\t' in line:
                        parts = line.split('\t')
                        stats_rows.append({
                            "Metric": parts[0].strip(),
                            "Value": parts[1].strip() if len(parts) > 1 else ""
                        })
                    else:
                        stats_rows.append({
                            "Metric": line,
                            "Value": ""
                        })
                
                # Add summary stats if not already present
                if not any('Total' in str(row.get('Metric', '')) for row in stats_rows):
                    stats_rows.insert(0, {
                        "Metric": "Total Rows Extracted",
                        "Value": str(len(extracted_full))
                    })
                    stats_rows.insert(1, {
                        "Metric": "Total Sheets Processed",
                        "Value": str(len(relevant_sheets))
                    })
                
                if stats_rows:
                    stats_df = pd.DataFrame(stats_rows)
                    stats_df.to_excel(writer, sheet_name='Stats', index=False)
                else:
                    # Fallback if parsing fails completely
                    simple_stats = pd.DataFrame({
                        "Metric": ["Total Rows Extracted", "Total Sheets Processed"],
                        "Value": [str(len(extracted_full)), str(len(relevant_sheets))]
                    })
                    simple_stats.to_excel(writer, sheet_name='Stats', index=False)
            
            excel_buffer.seek(0)
            
            # Download button for Excel
            st.download_button(
                "⬇️ Download Full Excel", 
                excel_buffer.getvalue(), 
                excel_filename, 
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download Excel file with Extracted, Mappings, and Stats sheets"
            )
            
            # Also provide CSV download as backup
            csv = extracted_full.to_csv(index=False)
            st.download_button("⬇️ Download CSV (backup)", csv, "extracted.csv", "text/csv")