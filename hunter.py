import pandas as pd, io, csv, string, re, sys
from collections import defaultdict
from datetime import datetime

def safe_print(*args, **kwargs):
    """Safely print without raising BrokenPipeError in Streamlit"""
    try:
        print(*args, **kwargs)
    except (BrokenPipeError, OSError):
        # Streamlit redirects stdout, pipe can close - ignore the error
        pass

def _find_cols(df, patterns, sheet):
    "Return list of (sheet, col_name) that match any pattern"
    hits = []
    for col_idx, col_name in enumerate(df.columns):
        name = str(col_name).lower()
        if any(p in name for p in patterns):
            hits.append((sheet, col_name))
    return hits

def _hunt_value(df, row_idx, col_refs, value_patterns):
    "Return first non-null value in any of the col_refs that matches value_patterns"
    for sh, col in col_refs:
        if sh != df.name: continue
        if col not in df.columns: continue
        val = str(df.iloc[row_idx, df.columns.get_loc(col)]).strip()
        if val=="" or val.lower()=="nan": continue
        if any(vp in val.lower() for vp in value_patterns):
            return val
    return None

def extract_data(all_sheets: dict[str, pd.DataFrame], mapping: dict):
    safe_print(f"🔍 Debug: extract_data called with {len(all_sheets)} sheets")
    safe_print(f"🔍 Debug: mapping = {mapping}")
    
    # Flatten mapping
    col_map = defaultdict(list)  # field -> [(sheet,col), …]
    for field, refs in mapping.items():
        safe_print(f"🔍 Debug: Processing field '{field}' with refs: {refs}")
        for r in refs:
            if "," in r:
                sheet_name, col_name = r.split(",", 1)
                col_name = col_name.strip()  # Remove any whitespace
                col_map[field].append((sheet_name, col_name))
                safe_print(f"🔍 Debug: Added mapping {field} -> sheet='{sheet_name}', column='{col_name}' (repr: {repr(col_name)})")
    
    safe_print(f"🔍 Debug: Final col_map = {dict(col_map)}")
    
    # Specifically check Relationship mapping
    if "Relationship To employee" in col_map:
        safe_print(f"🎯 RELATIONSHIP MAPPING DEBUG: {col_map['Relationship To employee']}")
    else:
        safe_print(f"⚠️ WARNING: 'Relationship To employee' not found in mapping!")

    # assemble master dataframe
    master = []
    safe_print(f"🔍 Debug: Processing {len(all_sheets)} sheets")
    for sh_name, df in all_sheets.items():
        safe_print(f"🔍 Debug: Processing sheet '{sh_name}' with {df.shape[0]} rows, {df.shape[1]} columns")
        safe_print(f"🔍 Debug: Columns in '{sh_name}': {list(df.columns)}")
        df = df.dropna(how="all").copy()
        df.name = sh_name
        safe_print(f"🔍 Debug: After dropna, '{sh_name}' has {len(df)} rows")
        
        # Define helper functions for this sheet
        def find_column(df, col_name):
            """Find column in DataFrame, handling whitespace and case differences"""
            col_name = str(col_name).strip()
            # Exact match first (highest priority)
            if col_name in df.columns:
                safe_print(f"🔍 find_column: Exact match found for '{col_name}'")
                return col_name
            # Try case-insensitive match
            for actual_col in df.columns:
                if str(actual_col).strip().lower() == col_name.lower():
                    safe_print(f"🔍 find_column: Case-insensitive match: '{col_name}' -> '{actual_col}'")
                    return actual_col
            # Try matching with whitespace normalization
            col_name_normalized = " ".join(col_name.split())
            for actual_col in df.columns:
                actual_normalized = " ".join(str(actual_col).split())
                if actual_normalized.lower() == col_name_normalized.lower():
                    safe_print(f"🔍 find_column: Whitespace-normalized match: '{col_name}' -> '{actual_col}'")
                    return actual_col
            safe_print(f"❌ find_column: No match found for '{col_name}'. Available columns: {list(df.columns)}")
            return None
        
        def split_full_name(full_name):
            """Split full name into first and last name"""
            if not full_name or not str(full_name).strip():
                return None, None
            
            full_name = str(full_name).strip()
            
            # Filter out NaN, nan, None, empty strings
            if not full_name or full_name.lower() in ["nan", "none", "null", ""]:
                return None, None
            
            # Handle comma-separated format: "Smith, John" or "Smith, John M"
            if "," in full_name:
                parts = [p.strip() for p in full_name.split(",")]
                if len(parts) >= 2:
                    return parts[1].strip(), parts[0].strip()  # Last name first, then first name
                elif len(parts) == 1:
                    return None, parts[0].strip()
            
            # Handle space-separated format: "John Smith" or "Mary Jane Doe"
            parts = full_name.split()
            if len(parts) >= 2:
                # Last word is last name, everything else is first name
                return " ".join(parts[:-1]).strip(), parts[-1].strip()
            elif len(parts) == 1:
                # Only one word - assume it's last name
                return None, parts[0].strip()
            
            return None, None
        
        def normalize_dob(dob_value):
            """Normalize DOB by removing timestamp and formatting as YYYY-MM-DD"""
            if not dob_value or str(dob_value).strip() == "" or str(dob_value).lower() == "nan":
                return ""
            
            dob_str = str(dob_value).strip()
            
            # Remove timestamp if present (format: "YYYY-MM-DD HH:MM:SS" or "YYYY-MM-DD 00:00:00")
            if " " in dob_str:
                dob_str = dob_str.split(" ")[0]
            
            # If it's in datetime format, try to parse and format
            try:
                # Try various date formats
                for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%m/%d/%y"]:
                    try:
                        dt = datetime.strptime(dob_str, fmt)
                        return dt.strftime("%Y-%m-%d")
                    except ValueError:
                        continue
            except:
                pass
            
            # If already in YYYY-MM-DD format, return as-is
            if len(dob_str) == 10 and dob_str[4] == "-" and dob_str[7] == "-":
                return dob_str
            
            # Return original if we can't parse it
            return dob_str
        
        def get_value_from_cols(col_refs):
            """Get value from columns, checking if it's a full name"""
            for sh, col in col_refs:
                if sh != sh_name: continue
                actual_col = find_column(df, col)
                if actual_col is None:
                    safe_print(f"🔍 Debug: Column '{col}' not found in sheet '{sh_name}'. Available columns: {list(df.columns)}")
                    continue
                val = df.iloc[row_idx, df.columns.get_loc(actual_col)]
                # Handle NaN values properly
                if pd.isna(val):
                    continue
                val = str(val).strip()
                # Filter out nan strings and empty values
                if val and val.lower() not in ["nan", "none", "null", ""]:
                    return val
            return None
        
        # Determine if this row is an employee or dependent row
        # Strategy: Check if Employee Name columns are filled (employee) vs Dependents/Relationship columns filled (dependent)
        
        for row_idx in range(len(df)):
            row_dict = {}
            
            # Check if this is likely an employee row (has Employee Name) or dependent row (has Dependents column)
            # Handle empty mappings (user selected "None")
            emp_cols = col_map.get("Employee Name", [])
            first_cols = col_map.get("First Name", [])
            last_cols = col_map.get("Last Name", [])
            rel_cols = col_map.get("Relationship To employee", [])
            
            # Check Employee Name column to see if this row has an employee
            employee_name_value = None
            for sh, col in emp_cols:
                if sh != sh_name: continue
                actual_col = find_column(df, col)
                if actual_col is None: continue
                emp_val_raw = df.iloc[row_idx, df.columns.get_loc(actual_col)]
                # Handle NaN values properly
                if pd.isna(emp_val_raw):
                    continue
                emp_val = str(emp_val_raw).strip()
                if emp_val and emp_val.lower() not in ["nan", "none", "null", ""]:
                    employee_name_value = emp_val
                    break
            
            # Check if there's a "Dependents" column or similar (unmapped column that might contain dependent names)
            dependents_col_value = None
            dependents_col_name = None
            relationship_value = None
            
            # First, check Relationship column if mapped
            for sh, col in rel_cols:
                if sh != sh_name: continue
                actual_col = find_column(df, col)
                if actual_col is None:
                    safe_print(f"🔍 Debug: Relationship column '{col}' not found in sheet '{sh_name}'. Available columns: {list(df.columns)}")
                    continue
                rel_val = str(df.iloc[row_idx, df.columns.get_loc(actual_col)]).strip()
                safe_print(f"🔍 Debug: Relationship mapped to '{col}' -> found column '{actual_col}' with value '{rel_val}' (row {row_idx})")
                if rel_val and rel_val.lower() not in ["nan", ""]:
                    relationship_value = rel_val
                    safe_print(f"✅ Debug: Using relationship value '{relationship_value}' from mapped column '{actual_col}'")
                    break
            
            # Look for unmapped columns that might contain dependent names
            for col_name in df.columns:
                actual_col = find_column(df, col_name)
                if actual_col is None: continue
                
                # Skip if already mapped
                is_mapped = False
                for mapped_field, mapped_cols in col_map.items():
                    for sh, mapped_col in mapped_cols:
                        if sh == sh_name and find_column(df, mapped_col) == actual_col:
                            is_mapped = True
                            break
                    if is_mapped:
                        break
                
                if not is_mapped:
                    col_lower = str(actual_col).lower()
                    col_value = str(df.iloc[row_idx, df.columns.get_loc(actual_col)]).strip()
                    
                    # Check if this looks like a "Dependents" column
                    if ("dependent" in col_lower or "spouse" in col_lower or "child" in col_lower) and col_value:
                        dependents_col_value = col_value
                        dependents_col_name = actual_col
                        break
            
            # Check if First Name/Last Name columns have values (might be a dependent row)
            first_name_val = get_value_from_cols(first_cols)
            last_name_val = get_value_from_cols(last_cols)
            has_name_values = bool(first_name_val or last_name_val)
            
            # Determine if this is an employee row or dependent row
            # First check relationship value - if it says Child/Spouse, treat as dependent regardless of Employee Name
            relationship_indicates_dependent = False
            relationship_is_job_title = False
            if relationship_value:
                rel_upper = str(relationship_value).upper().strip()
                # Check if it's a relationship term (dependent)
                relationship_indicates_dependent = any(keyword in rel_upper for keyword in [
                    "SPOUSE", "CHILD", "SON", "DAUGHTER", "WIFE", "HUSBAND", "DEPENDENT"
                ])
                # Check if it's a job title (employee) - job titles typically contain job-related keywords
                # and are longer/more descriptive than simple relationship terms
                job_title_keywords = ["MANAGER", "ASSISTANT", "ACCOUNTANT", "DIRECTOR", "COORDINATOR", 
                                      "SPECIALIST", "ANALYST", "SUPERVISOR", "EXECUTIVE", "OFFICER",
                                      "REPRESENTATIVE", "TECHNICIAN", "ADMINISTRATOR", "CARE", "PATIENT",
                                      "STAFF", "CLERK", "LEAD", "SENIOR", "JUNIOR", "PRINCIPAL"]
                relationship_is_job_title = any(keyword in rel_upper for keyword in job_title_keywords) or len(rel_upper) > 15
            
            # Employee row: 
            # 1. Has Employee Name filled AND relationship doesn't indicate dependent
            # 2. OR has name values (First/Last) AND relationship is a job title (not a relationship term)
            # If relationship says "Child" or "Spouse", it's a dependent even if Employee Name is filled
            is_employee_row = (
                (bool(employee_name_value) and not relationship_indicates_dependent) or
                (has_name_values and relationship_is_job_title and not relationship_indicates_dependent)
            )
            
            # **IMPORTANT**: Also check ALL unmapped columns for name-like values
            # This catches cases where dependent rows have names in unmapped columns
            name_in_unmapped = None
            if not has_name_values and not is_employee_row:
                # Priority: Check columns that look like name columns first
                name_priority_cols = []
                other_cols = []
                
                for col_name in df.columns:
                    actual_col = find_column(df, col_name)
                    if actual_col is None: continue
                    
                    # Skip if already mapped
                    is_mapped = False
                    for mapped_field, mapped_cols in col_map.items():
                        for sh, mapped_col in mapped_cols:
                            if sh == sh_name and find_column(df, mapped_col) == actual_col:
                                is_mapped = True
                                break
                        if is_mapped:
                            break
                    
                    if not is_mapped:
                        col_lower = str(actual_col).lower()
                        # Prioritize columns that look like name columns
                        if any(name_term in col_lower for name_term in ["first", "last", "name", "dependent"]):
                            name_priority_cols.append(actual_col)
                        else:
                            other_cols.append(actual_col)
                
                # Check priority columns first
                for actual_col in name_priority_cols + other_cols:
                    col_lower = str(actual_col).lower()
                    col_value = str(df.iloc[row_idx, df.columns.get_loc(actual_col)]).strip()
                    
                    # Check if this column contains a name-like value (has letters, spaces, might be a person's name)
                    if col_value and col_value.lower() not in ["nan", ""]:
                        # Check if it looks like a name (contains letters, possibly spaces, not just numbers)
                        if any(c.isalpha() for c in col_value) and len(col_value.strip()) > 2:
                            # Exclude obvious non-name columns
                            if not any(skip in col_lower for skip in ["zip", "code", "phone", "email", "address", "city", "state", "dob", "birth", "date", "id", "number", "ssn", "coverage", "plan"]):
                                name_in_unmapped = col_value
                                has_name_values = True  # Found a name in unmapped column
                                safe_print(f"🔍 Debug: Row {row_idx} - Found name-like value in unmapped column '{actual_col}': '{col_value}'")
                                break
            
            # Check if Relationship field is mapped - if it is, ONLY use mapped column, never override
            relationship_is_mapped = len(rel_cols) > 0
            
            # Only check unmapped columns if Relationship field is NOT mapped
            relationship_in_unmapped = None
            if not relationship_is_mapped:
                for col_name in df.columns:
                    actual_col = find_column(df, col_name)
                    if actual_col is None: continue
                    
                    # Skip if already mapped
                    is_mapped = False
                    for mapped_field, mapped_cols in col_map.items():
                        for sh, mapped_col in mapped_cols:
                            if sh == sh_name and find_column(df, mapped_col) == actual_col:
                                is_mapped = True
                                break
                        if is_mapped:
                            break
                    
                    if not is_mapped:
                        col_lower = str(actual_col).lower()
                        col_value = str(df.iloc[row_idx, df.columns.get_loc(actual_col)]).strip()
                        
                        # Check if this looks like a relationship column (contains spouse, child, etc.)
                        if ("relationship" in col_lower or "relation" in col_lower) and col_value:
                            relationship_in_unmapped = col_value
                            break
            
            # Dependent row detection: 
            # 1. If relationship value indicates dependent (Child, Spouse, etc.) - it's a dependent (highest priority)
            # 2. OR if Employee Name is empty BUT has name values or relationship info
            # IMPORTANT: If relationship is a job title, it's NOT a dependent (it's an employee)
            is_dependent_row = (
                relationship_indicates_dependent or  # Relationship value says it's a dependent (highest priority)
                (not is_employee_row and not relationship_is_job_title and (  # Not an employee row AND not a job title
                    bool(dependents_col_value) or  # Has Dependents column filled
                    (bool(relationship_value) and not relationship_is_job_title) or  # Has Relationship column filled (mapped) AND it's not a job title
                    (not relationship_is_mapped and bool(relationship_in_unmapped)) or  # Has Relationship column filled (unmapped) - only if not mapped
                    (has_name_values and relationship_value and not relationship_is_job_title) or  # Has name AND relationship (mapped) AND it's not a job title
                    (has_name_values and not relationship_is_mapped and relationship_in_unmapped) or  # Has name AND relationship (unmapped) - only if not mapped
                    (has_name_values and not relationship_value and not relationship_is_job_title)  # Just has name values (First/Last Name) without Employee Name = likely dependent
                ))
            )
            
            # Use unmapped relationship ONLY if Relationship field is NOT mapped
            if not relationship_is_mapped and relationship_in_unmapped:
                relationship_value = relationship_in_unmapped
            
            safe_print(f"🔍 Debug: Row {row_idx} - Employee Name: '{employee_name_value}', Dependents: '{dependents_col_value}', Relationship: '{relationship_value}'")
            safe_print(f"   → Is Employee Row: {is_employee_row}, Is Dependent Row: {is_dependent_row}")
            
            first, last = None, None
            
            if is_employee_row:
                # Priority 1: If separate First Name and Last Name columns exist and have values, use them directly
                first_val = get_value_from_cols(first_cols)
                last_val = get_value_from_cols(last_cols)
                
                # Check if First Name and Last Name columns have actual separate values (not full names)
                has_separate_name_columns = False
                if first_val and last_val:
                    first_clean = str(first_val).strip().lower()
                    last_clean = str(last_val).strip().lower()
                    if (first_clean not in ["nan", "none", "null", ""] and 
                        last_clean not in ["nan", "none", "null", ""]):
                        # Check if First Name column doesn't contain a full name (no comma, single word or proper first name pattern)
                        # If First Name has multiple words and Last Name is empty or single word, it might be a full name
                        first_parts = first_clean.split()
                        last_parts = last_clean.split()
                        # If First Name has 1-2 words and Last Name has 1 word, likely separate columns
                        # If First Name has 2+ words and Last Name is empty/weak, First Name might be a full name
                        if len(first_parts) <= 2 and len(last_parts) == 1:
                            has_separate_name_columns = True
                            first = first_val
                            last = last_val
                            safe_print(f"🔍 Debug: Employee row {row_idx} - Using separate First Name='{first}' and Last Name='{last}' columns")
                
                # Priority 2: If we have Employee Name column (combined name), split it
                if not has_separate_name_columns and employee_name_value:
                    first, last = split_full_name(employee_name_value)
                    safe_print(f"🔍 Debug: Employee row {row_idx} - Split Employee Name '{employee_name_value}' → first='{first}', last='{last}'")
                
                # Priority 3: Fallback - use First Name/Last Name even if they might be full names (split if needed)
                if not first or not last:
                    if first_val and first_val.lower() not in ["nan", "none", "null", ""]:
                        # Check if First Name contains a full name (has multiple words)
                        first_parts = str(first_val).split()
                        if len(first_parts) > 1:
                            # Might be a full name, try to split it
                            temp_first, temp_last = split_full_name(first_val)
                            if temp_first and temp_last:
                                first = temp_first
                                last = temp_last if not last_val else last_val
                                safe_print(f"🔍 Debug: Employee row {row_idx} - Split First Name column '{first_val}' → first='{first}', last='{last}'")
                            else:
                                first = first_val
                        else:
                            first = first_val
                    if last_val and last_val.lower() not in ["nan", "none", "null", ""]:
                        last = last_val
                
                # Filter out "nan" strings from first and last names
                if first and str(first).lower() in ["nan", "none", "null"]:
                    first = None
                if last and str(last).lower() in ["nan", "none", "null"]:
                    last = None
                
                if not first and not last:
                    safe_print(f"🔍 Debug: Skipping employee row {row_idx} - no name found")
                    continue
                
                row_dict["First Name"] = str(first).strip() if first else ""
                row_dict["Last Name"] = str(last).strip() if last else ""
                # Extract relationship value directly from mapped column - no substitution, no defaults
                if relationship_value:
                    row_dict["Relationship To employee"] = relationship_value
                    safe_print(f"🔍 Debug: Employee row {row_idx} - Using relationship value '{relationship_value}' directly from mapped column")
                else:
                    row_dict["Relationship To employee"] = ""  # Empty if no value found, don't invent values
                    safe_print(f"🔍 Debug: Employee row {row_idx} - No relationship value found, leaving empty")
                row_dict["Dependent (Y/N)"] = "N"
                
            elif is_dependent_row:
                # Extract dependent - try multiple sources with correct priority
                # Priority 1: If separate First Name and Last Name columns exist and have values, use them directly
                first_val = get_value_from_cols(first_cols)
                last_val = get_value_from_cols(last_cols)
                
                has_separate_name_columns = False
                if first_val and last_val:
                    first_clean = str(first_val).strip().lower()
                    last_clean = str(last_val).strip().lower()
                    if (first_clean not in ["nan", "none", "null", ""] and 
                        last_clean not in ["nan", "none", "null", ""]):
                        # Check if these look like separate columns
                        first_parts = first_clean.split()
                        last_parts = last_clean.split()
                        if len(first_parts) <= 2 and len(last_parts) == 1:
                            has_separate_name_columns = True
                            first = first_val
                            last = last_val
                            safe_print(f"🔍 Debug: Dependent row {row_idx} - Using separate First Name='{first}' and Last Name='{last}' columns")
                
                # Priority 2: Dependents column (combined name)
                if not has_separate_name_columns and dependents_col_value and dependents_col_value.lower() not in ["nan", "none", "null", ""]:
                    dep_name = str(dependents_col_value).strip()
                    # Remove relationship info if present in parentheses
                    if "(" in dep_name:
                        dep_name = dep_name.split("(")[0].strip()
                    first, last = split_full_name(dep_name)
                    safe_print(f"🔍 Debug: Dependent row {row_idx} - Split Dependents column '{dep_name}' → first='{first}', last='{last}'")
                
                # Priority 3: Name in unmapped column (combined name)
                if (not first or not last) and name_in_unmapped:
                    if str(name_in_unmapped).lower() not in ["nan", "none", "null", ""]:
                        first, last = split_full_name(name_in_unmapped)
                        safe_print(f"🔍 Debug: Dependent row {row_idx} - Split unmapped column '{name_in_unmapped}' → first='{first}', last='{last}'")
                
                # Priority 4: Fallback - First Name/Last Name columns, split if they contain full names
                if not first or not last:
                    if first_val and first_val.lower() not in ["nan", "none", "null", ""]:
                        if " " in first_val or "," in first_val:
                            dep_first, dep_last = split_full_name(first_val)
                            if dep_first and dep_first.lower() not in ["nan", "none", "null"]:
                                first = dep_first
                            if dep_last and dep_last.lower() not in ["nan", "none", "null"]:
                                last = dep_last
                            safe_print(f"🔍 Debug: Dependent row {row_idx} - Split First Name column '{first_val}' → first='{first}', last='{last}'")
                        else:
                            first = first_val
                    
                    if last_val and last_val.lower() not in ["nan", "none", "null", ""]:
                        if " " in last_val or "," in last_val:
                            dep_first, dep_last = split_full_name(last_val)
                            if dep_first and dep_first.lower() not in ["nan", "none", "null"]:
                                first = dep_first
                            if dep_last and dep_last.lower() not in ["nan", "none", "null"]:
                                last = dep_last
                        else:
                            if not last:
                                last = last_val
                    
                    if first or last:
                        safe_print(f"🔍 Debug: Dependent row {row_idx} - extracted from First/Last Name columns: first='{first}', last='{last}'")
                
                # Filter out "nan" strings from first and last names
                if first and str(first).lower() in ["nan", "none", "null"]:
                    first = None
                if last and str(last).lower() in ["nan", "none", "null"]:
                    last = None
                
                if not first and not last:
                    safe_print(f"🔍 Debug: Skipping dependent row {row_idx} - no name found")
                    continue
                
                row_dict["First Name"] = str(first).strip() if first else ""
                row_dict["Last Name"] = str(last).strip() if last else ""
                
                # Extract relationship from Relationship column - use value as-is, no normalization
                safe_print(f"🔍 Debug: Dependent row {row_idx} - Setting relationship from relationship_value='{relationship_value}'")
                if relationship_value:
                    row_dict["Relationship To employee"] = relationship_value
                    safe_print(f"🔍 Debug: Using relationship value as-is: '{relationship_value}'")
                else:
                    row_dict["Relationship To employee"] = ""  # Empty if no value found, don't invent values
                    safe_print(f"🔍 Debug: No relationship value found, leaving empty")
                
                row_dict["Dependent (Y/N)"] = "Y"
                
            else:
                # Neither employee nor dependent row - skip
                safe_print(f"🔍 Debug: Skipping row {row_idx} - not identified as employee or dependent")
                continue

            # ---- Extract other fields ------------------------------------------
            for field in ["DOB","Gender","Medical Coverage","Medical Plan Name",
                          "Dental Coverage","Dental Plan Name",
                          "Vision Coverage","Vision Plan Name",
                          "COBRA Participation (Y/N)"]:
                # Check if field is mapped (not empty list)
                field_cols = col_map.get(field, [])
                if not field_cols or len(field_cols) == 0:
                    # Field is unmapped - leave empty
                    row_dict[field] = ""
                    continue
                
                vals = []
                for sh,col in field_cols:
                    if sh!=sh_name: continue
                    actual_col = find_column(df, col)
                    if actual_col is None:
                        continue
                    v = str(df.iloc[row_idx, df.columns.get_loc(actual_col)]).strip()
                    if v and v.lower()!="nan":
                        vals.append(v)
                # Normalize DOB to remove timestamps
                if field == "DOB":
                    row_dict[field] = normalize_dob(vals[0]) if vals else ""
                else:
                    row_dict[field] = vals[0] if vals else ""

            # For dependent rows, extract DOB from Date Of Birth column if available
            if is_dependent_row:
                # Look for Date Of Birth column (might be unmapped)
                for col_name in df.columns:
                    actual_col = find_column(df, col_name)
                    if actual_col is None:
                        continue
                    col_lower = str(actual_col).lower()
                    if "birth" in col_lower or "dob" in col_lower:
                        dob_val = str(df.iloc[row_idx, df.columns.get_loc(actual_col)]).strip()
                        if dob_val and dob_val.lower() not in ["nan", ""]:
                            row_dict["DOB"] = normalize_dob(dob_val)
                            break

            row_dict["__sheet__"] = sh_name
            row_dict["__original_row_idx__"] = row_idx  # Preserve original row order for grouping
            row_dict["__sheet_name__"] = sh_name
            
            # Add the record (employee or dependent)
            if is_employee_row:
                # Add employee record
                master.append(row_dict.copy())
                safe_print(f"🔍 Debug: Added employee row {row_idx} to master: {row_dict.get('First Name', '')} {row_dict.get('Last Name', '')}")
                
                # Check if employee row also has a dependent in Dependents column (same row)
                if dependents_col_value:
                    # Parse dependent name from Dependents column (might contain relationship info)
                    dep_name = dependents_col_value
                    
                    # Extract relationship from Dependents column if present (e.g., "... (Relationship: WIFE, ...)")
                    extracted_relationship = None
                    extracted_dob = None
                    
                    if "(" in dep_name:
                        # Try to extract relationship from parentheses
                        paren_content = dep_name[dep_name.index("(")+1:dep_name.rindex(")")]
                        dep_name = dep_name.split("(")[0].strip()
                        
                        # Look for "Relationship:" pattern
                        if "relationship" in paren_content.lower():
                            rel_match = None
                            if "relationship:" in paren_content.lower():
                                parts = paren_content.split("relationship:")
                                if len(parts) > 1:
                                    rel_part = parts[1].split(",")[0].strip()
                                    extracted_relationship = rel_part
                            elif "relationship" in paren_content.lower():
                                # Try other patterns
                                for word in ["WIFE", "SPOUSE", "SON", "DAUGHTER", "CHILD"]:
                                    if word in paren_content.upper():
                                        extracted_relationship = word
                                        break
                        
                        # Look for DOB in parentheses (Date Of Birth: or DOB:)
                        if "date of birth" in paren_content.lower():
                            # Extract after "Date Of Birth:" or "Date of Birth:"
                            dob_parts = paren_content.split("Date Of Birth:")
                            if len(dob_parts) < 2:
                                dob_parts = paren_content.split("Date of Birth:")
                            if len(dob_parts) < 2:
                                dob_parts = paren_content.split("date of birth:")
                            if len(dob_parts) >= 2:
                                dob_part = dob_parts[1].strip()
                                if "," in dob_part:
                                    dob_part = dob_part.split(",")[0].strip()
                                extracted_dob = dob_part
                        elif "dob:" in paren_content.lower():
                            dob_parts = paren_content.split("DOB:")
                            if len(dob_parts) < 2:
                                dob_parts = paren_content.split("dob:")
                            if len(dob_parts) >= 2:
                                dob_part = dob_parts[1].strip()
                                if "," in dob_part:
                                    dob_part = dob_part.split(",")[0].strip()
                                extracted_dob = dob_part
                    
                    dep_first, dep_last = split_full_name(dep_name)
                    if dep_first or dep_last:
                        # Create dependent record from same row
                        dependent_record = row_dict.copy()
                        dependent_record["First Name"] = dep_first or ""
                        dependent_record["Last Name"] = dep_last or ""
                        
                        # Extract relationship - use value from mapped column as-is, no normalization
                        if extracted_relationship:
                            dependent_record["Relationship To employee"] = extracted_relationship
                        elif relationship_value:
                            dependent_record["Relationship To employee"] = relationship_value
                        else:
                            dependent_record["Relationship To employee"] = ""  # Empty if no value, don't invent values
                        
                        # Set DOB if extracted (normalize to remove timestamp)
                        if extracted_dob:
                            dependent_record["DOB"] = normalize_dob(extracted_dob)
                        
                        dependent_record["Dependent (Y/N)"] = "Y"
                        dependent_record["__original_row_idx__"] = row_idx  # Same row as employee
                        master.append(dependent_record)
                        safe_print(f"🔍 Debug: Added dependent from same row {row_idx}: {dep_first} {dep_last} (Relationship: {dependent_record.get('Relationship To employee', 'Unknown')}) (from Dependents column)")
            
            elif is_dependent_row:
                # Add dependent record (will be grouped with employee later)
                master.append(row_dict.copy())
                safe_print(f"🔍 Debug: Added dependent row {row_idx} to master: {row_dict.get('First Name', '')} {row_dict.get('Last Name', '')} (Relationship: {row_dict.get('Relationship To employee', '')})")

    # Group employees and dependents based on row proximity and last name
    safe_print(f"🔄 Grouping employees and dependents...")
    grouped_master = group_employees_and_dependents(master)
    
    # Reorder grouped records to maintain original Excel row order
    # Sort by original row index to preserve Excel order
    grouped_master_sorted = sorted(grouped_master, key=lambda x: (
        x.get("__sheet_name__", ""),
        x.get("__original_row_idx__", 999999)
    ))
    
    extracted = pd.DataFrame(grouped_master_sorted)
    safe_print(f"🔍 Debug: Created DataFrame with {len(grouped_master_sorted)} rows")
    safe_print(f"🔍 Debug: DataFrame columns: {list(extracted.columns) if len(grouped_master_sorted) > 0 else 'No data'}")
    if len(grouped_master_sorted) > 0:
        safe_print(f"🔍 Debug: First row: {grouped_master_sorted[0]}")
    
    # Remove internal grouping columns before returning
    if "__original_row_idx__" in extracted.columns:
        extracted = extracted.drop(columns=["__original_row_idx__"])
    if "__sheet_name__" in extracted.columns:
        extracted = extracted.drop(columns=["__sheet_name__"])
    
    # Convert datetime objects to strings to prevent PyArrow serialization errors in Streamlit
    # Convert datetime64 columns first
    datetime_cols = extracted.select_dtypes(include=['datetime64']).columns
    for col in datetime_cols:
        extracted[col] = extracted[col].apply(lambda x: str(x) if pd.notna(x) else "")
    
    # Then check object columns for datetime objects
    for col in extracted.columns:
        if extracted[col].dtype == 'object':
            # Check if column contains datetime objects
            non_null_values = extracted[col].dropna()
            if len(non_null_values) > 0:
                try:
                    sample_val = non_null_values.iloc[0]
                    if isinstance(sample_val, pd.Timestamp) or (hasattr(sample_val, '__class__') and 'datetime' in str(type(sample_val)).lower()):
                        extracted[col] = extracted[col].apply(lambda x: str(x) if pd.notna(x) and x is not None else "")
                except (IndexError, AttributeError):
                    pass
    
    # Reorder columns: Family Group first, then names, then Gender followed by Relationship/Dependent fields
    if "Family Group" in extracted.columns:
        cols = extracted.columns.tolist()
        # Remove Family Group from its current position
        cols.remove("Family Group")
        # Insert Family Group before First Name (or at beginning if First Name doesn't exist)
        if "First Name" in cols:
            first_name_idx = cols.index("First Name")
            cols.insert(first_name_idx, "Family Group")
        else:
            cols.insert(0, "Family Group")
    
    # Reorder: Move Relationship To employee, Dependent Of Employee Row, and Dependent (Y/N) next to Gender
    cols = extracted.columns.tolist()
    
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
            if field in extracted.columns:
                cols.insert(insert_pos, field)
                insert_pos += 1
    
    extracted = extracted[cols]
    
    stats = produce_stats(all_sheets, extracted)
    return extracted, stats

def group_employees_and_dependents(master_list):
    """
    Group employees and dependents based on:
    1. Proximity (same row or succeeding rows) - HIGHEST PRIORITY
    2. Last name matching - SECONDARY
    Preserves original Excel row order in final output.
    Assigns Family Group numbers to employees and their dependents.
    """
    safe_print(f"🔄 Grouping {len(master_list)} records...")
    
    if not master_list:
        return master_list
    
    # Sort records by sheet and original row index to maintain Excel order
    sorted_records = sorted(enumerate(master_list), key=lambda x: (
        x[1].get("__sheet_name__", ""),
        x[1].get("__original_row_idx__", 999999)
    ))
    
    grouped_records = []
    used_indices = set()
    family_group_counter = 1  # Start family group numbering at 1
    
    # Process records in order - employees first, then dependents
    # First pass: Process all employees and assign Family Group numbers
    employee_family_groups = {}  # Map (sheet, row_idx) -> family_group_number
    employee_records_map = {}  # Map (sheet, row_idx) -> employee_record
    
    for idx, record in sorted_records:
        relationship = str(record.get("Relationship To employee", "")).lower().strip()
        is_dependent = (
            relationship in ["spouse", "child", "son", "daughter", "wife", "husband", "dependent"] or
            any(k in relationship for k in ["spouse", "child", "dependent"])
        )
        
        if not is_dependent:
            # This is an employee - assign family group number
            employee_record = dict(record)
            employee_record["Family Group"] = family_group_counter
            employee_key = (record.get("__sheet_name__", ""), record.get("__original_row_idx__", -1))
            employee_family_groups[employee_key] = family_group_counter
            employee_records_map[employee_key] = employee_record
            grouped_records.append(employee_record)
            used_indices.add(idx)
            emp_name = f"{record.get('First Name', '')} {record.get('Last Name', '')}"
            safe_print(f"   👤 Added employee: {emp_name} (Family Group: {family_group_counter}, row {record.get('__original_row_idx__', -1)})")
            family_group_counter += 1
    
    # Second pass: Process dependents and link them to employees
    for idx, record in sorted_records:
        if idx in used_indices:
            continue
        
        relationship = str(record.get("Relationship To employee", "")).lower().strip()
        
        # Check if this is a dependent
        is_dependent = (
            relationship in ["spouse", "child", "son", "daughter", "wife", "husband", "dependent"] or
            any(k in relationship for k in ["spouse", "child", "dependent"])
        )
        
        if is_dependent:
            # This is a dependent - find the employee it belongs to
            employee_found = False
            record_sheet = record.get("__sheet_name__", "")
            record_row = record.get("__original_row_idx__", -1)
            record_last_name = str(record.get("Last Name", "")).lower().strip()
            dependent_family_group = None
            
            # Strategy: Find the employee this dependent belongs to
            # Priority 1: Last Name matching (HIGHEST)
            # Priority 2: Proximity (same row or succeeding rows) (SECONDARY)
            best_employee_key = None
            best_employee_record = None
            min_distance = 999999
            
            # Search through all employees that were processed in first pass
            candidates_same_lastname = []  # Employees with matching last name
            candidates_by_proximity = []   # Employees by proximity
            
            for emp_key, emp_family_group in employee_family_groups.items():
                # Get employee record from map
                emp_record = employee_records_map.get(emp_key)
                if not emp_record:
                    continue
                
                check_sheet = emp_record.get("__sheet_name__", "")
                check_row = emp_record.get("__original_row_idx__", -1)
                
                # Only consider employees from same sheet, before the dependent row
                if check_sheet != record_sheet or check_row >= record_row:
                    continue
                
                check_last_name = str(emp_record.get("Last Name", "")).lower().strip()
                distance = record_row - check_row  # Positive because check_row < record_row
                
                # Priority 1: Same last name (HIGHEST PRIORITY)
                if check_last_name == record_last_name and distance <= 5:
                    candidates_same_lastname.append((emp_key, emp_record, distance))
                
                # Priority 2: Proximity - within 5 rows (SECONDARY PRIORITY)
                elif distance <= 5:
                    candidates_by_proximity.append((emp_key, emp_record, distance))
            
            # Select best employee: First from same last name, then by proximity
            if candidates_same_lastname:
                # Same last name - pick closest (smallest distance)
                best_employee_key, best_employee_record, min_distance = min(candidates_same_lastname, key=lambda x: x[2])
                safe_print(f"   🎯 Matched by LAST NAME: {record.get('Last Name', '')} (distance: {min_distance})")
            elif candidates_by_proximity:
                # No last name match - pick closest by proximity
                best_employee_key, best_employee_record, min_distance = min(candidates_by_proximity, key=lambda x: x[2])
                safe_print(f"   🎯 Matched by PROXIMITY only (different last name, distance: {min_distance})")
            
            if best_employee_key is not None and best_employee_record is not None:
                # Found employee - get the employee's family group number
                dependent_family_group = employee_family_groups.get(best_employee_key)
                
                if dependent_family_group is not None:
                    # Find where the employee is in grouped_records to insert dependent after it
                    employee_in_grouped = None
                    for g_idx, grouped_record in enumerate(grouped_records):
                        if (grouped_record.get("__sheet_name__") == best_employee_key[0] and 
                            grouped_record.get("__original_row_idx__") == best_employee_key[1] and
                            grouped_record.get("Family Group") == dependent_family_group):
                            employee_in_grouped = g_idx
                            break
                    
                    if employee_in_grouped is not None:
                        # Insert dependent after employee and any existing dependents from same family
                        insert_pos = employee_in_grouped + 1
                        
                        # Insert after employee and any existing dependents from same family group
                        while (insert_pos < len(grouped_records) and
                               grouped_records[insert_pos].get("Family Group") == dependent_family_group):
                            insert_pos += 1
                        
                        # Create dependent record with the same Family Group number as the employee
                        dependent_record = dict(record)
                        dependent_record["Family Group"] = dependent_family_group  # Same number as employee!
                        grouped_records.insert(insert_pos, dependent_record)
                        used_indices.add(idx)
                        employee_found = True
                        emp_name = f"{best_employee_record.get('First Name', '')} {best_employee_record.get('Last Name', '')}"
                        dep_name = f"{record.get('First Name', '')} {record.get('Last Name', '')}"
                        row_info = "same row" if min_distance == 0 else f"{min_distance} rows after employee"
                        safe_print(f"   ✅ Grouped dependent {dep_name} with employee {emp_name} (Family Group: {dependent_family_group}) ({row_info})")
                    else:
                        safe_print(f"   ⚠️ Employee found but not in grouped_records: {best_employee_record.get('First Name', '')} {best_employee_record.get('Last Name', '')}")
                else:
                    safe_print(f"   ⚠️ Employee found but no Family Group number: {best_employee_record.get('First Name', '')} {best_employee_record.get('Last Name', '')}")
            
            if not employee_found:
                # No employee found - assign new family group number (standalone dependent becomes a new family)
                standalone_record = dict(record)
                standalone_record["Family Group"] = family_group_counter
                family_group_counter += 1
                grouped_records.append(standalone_record)
                used_indices.add(idx)
                safe_print(f"   ⚠️ Unmatched dependent: {record.get('First Name', '')} {record.get('Last Name', '')} (Family Group: {standalone_record['Family Group']})")
    
    safe_print(f"👨‍👩‍👧‍👦 Grouped {len(grouped_records)} records into {family_group_counter - 1} family groups")
    return grouped_records

def produce_stats(all_sheets, extracted):
    """Produce statistics in markdown table format."""
    out = io.StringIO()
    
    # Markdown format
    out.write("## Extraction Statistics\n\n")
    
    # Summary table
    out.write("### Summary\n\n")
    out.write("| Metric | Value |\n")
    out.write("|--------|-------|\n")
    out.write(f"| **Filename (Excel)** | Uploaded file |\n")
    out.write(f"| **Total Sheets** | {len(all_sheets)} |\n")
    out.write(f"| **Total Extracted Records** | {len(extracted)} |\n")
    
    # Count employees and dependents if available
    if 'Relationship To employee' in extracted.columns:
        employees = len(extracted[extracted['Relationship To employee'].str.lower().str.strip() == 'employee'])
        dependents = len(extracted) - employees
        out.write(f"| **Employees** | {employees} |\n")
        out.write(f"| **Dependents** | {dependents} |\n")
    
    out.write("\n### Sheet Details\n\n")
    out.write("| Sheet Name | Rows | Columns |\n")
    out.write("|------------|------|---------|\n")
    for name, df in all_sheets.items():
        out.write(f"| {name} | {df.shape[0]} | {df.shape[1]} |\n")
    
    return out.getvalue()

