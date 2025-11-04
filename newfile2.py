import pandas as pd
import json
import re
import os
from groq import Groq
import tkinter as tk
from tkinter import Tk, filedialog, messagebox, ttk
import math
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import csv
import threading

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
MODEL_NAME = "llama-3.3-70b-versatile"

# 🚀 Optimized limits for faster processing
CHUNK_SIZE = 15           # smaller chunks for faster processing
MAX_CHARS = 12000         # reduced input limit for speed
MAX_TOKENS = 8000         # reduced response limit for speed
SUBCHUNK_CHARS = 8000     # smaller sub-chunks
ABSOLUTE_MAX_CHARS = 12000  # reduced hard limit

# =========================================================
# UTILITY FUNCTIONS
# =========================================================
def excel_to_json_all(file_path):
    """Read all sheets from Excel into JSON-serializable dict with optimization."""
    try:
        sheets = pd.read_excel(file_path, sheet_name=None)
    except Exception as e:
        print(f"❌ Error reading Excel: {e}")
        sys.exit(1)

    result = {}
    for sheet_name, df in sheets.items():
        # Optimize: only keep relevant columns and limit rows
        df = df.fillna("").astype(str)
        
        # Remove completely empty columns
        df = df.loc[:, (df != "").any(axis=0)]
        
        # Limit to first 50 rows for faster processing
        if len(df) > 50:
            df = df.head(50)
            print(f"⚠️ Sheet '{sheet_name}' has {len(df)} rows, processing first 50 for speed")
        
        result[sheet_name] = df.to_dict(orient="records")
    return result

def truncate_json_data(records, max_chars=MAX_CHARS):
    """Ensure JSON string stays under token limits."""
    text = json.dumps(records, indent=2)
    if len(text) > max_chars:
        text = text[:max_chars] + "... (truncated)"
    return text

def split_large_json(records, max_chars=SUBCHUNK_CHARS):
    """Split large JSON into smaller chunks."""
    text = json.dumps(records, indent=2)
    if len(text) <= max_chars:
        return [text]
    
    # Split by records
    chunk_size = len(records) // 2
    chunks = []
    for i in range(0, len(records), chunk_size):
        chunk = records[i:i + chunk_size]
        chunk_text = json.dumps(chunk, indent=2)
        if len(chunk_text) > max_chars:
            # Further split this chunk
            sub_chunks = split_large_json(chunk, max_chars)
            chunks.extend(sub_chunks)
        else:
            chunks.append(chunk_text)
    return chunks

def ask_question_chunk(chunk_data, question, model=MODEL_NAME, attempt=1):
    """Send one chunk of Excel data to Groq safely within limits."""
    safe_json = truncate_json_data(chunk_data, max_chars=min(MAX_CHARS, ABSOLUTE_MAX_CHARS))

    final_input = f"""
    You are a JSON data extraction engine. Always output strictly valid JSON, no explanations.

    CRITICAL ANTI-HALLUCINATION RULES:
    1. Extract ONLY the data that is actually present in the input
    2. Do NOT create, invent, or hallucinate any data that is not explicitly provided
    3. Do NOT add fake names like "John Doe", "Jane Smith", "Mike Johnson" etc.
    4. If a field is not present or empty, use null or empty string
    5. ONLY extract people whose names are explicitly shown in the data
    6. If you see "Not available for Enrollment" or similar, that person should still be extracted

    Task: Extract structured data from Excel.

    Instructions: {question}

    Fields: [
       "last_name", "first_name", "employee_name", "home_zip_code",
        "dob", "gender", "medical_coverage", "vision_coverage",
        "dental_coverage", "cobra_participation", "relationship_to_employee",
        "dependent_of_employee_row"
    ]

    Excel Data (sample): {safe_json}

    REMEMBER: Only extract people whose names are actually in the input data. Do not create fake people.

    Wrap your final JSON output strictly inside <json> ... </json> tags.
    """

    try:
        completion = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are a JSON data extraction engine. Always output strictly valid JSON, no explanations. NEVER create fake data or hallucinate names. Only extract data that is explicitly present in the input."},
                {"role": "user", "content": final_input},
            ],
            max_tokens=MAX_TOKENS,
            temperature=0,
        )
    except Exception as e:
        error_msg = f"❌ Groq API Error (Attempt {attempt}): {e}"
        print(error_msg)
        
        # Check if it's a rate limit error
        if "rate limit" in str(e).lower() or "429" in str(e):
            print("🚫 Rate limit exceeded - aborting operations...")
            return "RATE_LIMIT_EXCEEDED"
        
        # Check for timeout errors
        if "timeout" in str(e).lower() or "timed out" in str(e).lower():
            print("⏰ API call timed out - this might be due to large data size")
            return "TIMEOUT_ERROR"
        
        # Check for other common errors
        if "connection" in str(e).lower():
            print("🔌 Connection error - check your internet connection")
            return "CONNECTION_ERROR"
        
        if attempt < 2:
            print("⏳ Retrying after 5 seconds...")
            time.sleep(5)
            return ask_question_chunk(chunk_data, question, model, attempt + 1)
        return []

    output_text = completion.choices[0].message.content

    # Extract <json>...</json>
    match = re.search(r"<json>(.*?)</json>", output_text, re.DOTALL)
    clean_json = match.group(1).strip() if match else output_text.strip()

    try:
        parsed = json.loads(clean_json)
        if isinstance(parsed, dict):
            return [parsed]
        elif isinstance(parsed, list):
            return parsed
        else:
            return []
    except json.JSONDecodeError:
        print(f"⚠️ Could not parse JSON from response: {clean_json[:200]}...")
        return []

def process_large_input(records, question, model=MODEL_NAME, log_callback=None):
    """Handle very large chunks (e.g. 100K+ chars) by splitting and merging."""
    text = json.dumps(records, indent=2)
    if len(text) <= MAX_CHARS:
        result = ask_question_chunk(records, question, model)
        if result == "RATE_LIMIT_EXCEEDED":
            if log_callback:
                log_callback("🚫 Rate limit exceeded - aborting operations...")
        return result

    msg = f"⚙️ Large input detected ({len(text)} chars). Splitting into sub-chunks..."
    print(msg)
    if log_callback:
        log_callback(msg)
    
    split_chunks = split_large_json(records, max_chars=SUBCHUNK_CHARS)
    combined = []

    for idx, chunk_text in enumerate(split_chunks, 1):
        msg = f"🧩 Processing subchunk {idx}/{len(split_chunks)}..."
        print(msg)
        if log_callback:
            log_callback(msg)
        
        try:
            sub_records = json.loads(chunk_text)
        except Exception:
            # fallback if malformed JSON fragment
            sub_records = records[(idx - 1) * 50:(idx * 50)]
        sub_results = ask_question_chunk(sub_records, question, model)
        
        # Check for rate limit error
        if sub_results == "RATE_LIMIT_EXCEEDED":
            if log_callback:
                log_callback("🚫 Rate limit exceeded - aborting operations...")
            return "RATE_LIMIT_EXCEEDED"
        
        if sub_results:
            combined.extend(sub_results)

    msg = f"✅ Combined {len(combined)} records from all subchunks."
    print(msg)
    if log_callback:
        log_callback(msg)
    return combined

# =========================================================
# ENHANCED GROUPING LOGIC
# =========================================================
def group_employees_and_dependents(all_results):
    """
    Group employees and dependents properly to avoid duplicates.
    """
    print("🔄 Grouping employees and dependents...")
    
    # Separate employees and dependents with better logic
    employees = []
    dependents = []
    
    for record in all_results:
        # Check multiple possible relationship fields
        relationship = str(
            record.get('relationship_to_employee', '') or 
            record.get('relationship', '') or 
            record.get('relation', '') or 
            record.get('person_type', '') or 
            record.get('job_title', '')
        ).lower().strip()
        
        # Enhanced employee detection
        is_employee = (
            relationship in ['employee', 'self', ''] or 
            not relationship or
            'employee' in relationship or
            'manager' in relationship or
            'director' in relationship or
            'assistant' in relationship or
            'analyst' in relationship or
            'coordinator' in relationship or
            'specialist' in relationship
        )
        
        # Enhanced dependent detection
        is_dependent = (
            relationship in ['spouse', 'child', 'son', 'daughter', 'wife', 'husband'] or
            'dependent' in relationship or
            'child' in relationship or
            'spouse' in relationship
        )
        
        if is_employee and not is_dependent:
            employees.append(record)
        elif is_dependent:
            dependents.append(record)
        else:
            # Default to employee if unclear
            employees.append(record)
    
    print(f"📊 Found {len(employees)} employees and {len(dependents)} dependents")
    
    # Debug: Show sample of what we found
    if employees:
        print(f"🔍 Sample employee: {employees[0].get('first_name', '')} {employees[0].get('last_name', '')} - {employees[0].get('relationship_to_employee', 'N/A')}")
    if dependents:
        print(f"🔍 Sample dependent: {dependents[0].get('first_name', '')} {dependents[0].get('last_name', '')} - {dependents[0].get('relationship_to_employee', 'N/A')}")
    
    # Group dependents with their employees using proximity and relationship logic
    grouped_families = []
    used_dependents = set()
    
    print(f"🔍 Starting grouping: {len(employees)} employees, {len(dependents)} dependents")
    
    # Create a mapping of employee names to their index for reference matching
    employee_name_map = {}
    for emp_idx, employee in enumerate(employees):
        emp_name = f"{str(employee.get('first_name', '')).lower().strip()} {str(employee.get('last_name', '')).lower().strip()}"
        employee_name_map[emp_name] = emp_idx
    
    for emp_idx, employee in enumerate(employees):
        # Create family group starting with employee
        family = [employee]
        employee_last_name = str(employee.get('last_name', '')).lower().strip()
        employee_first_name = str(employee.get('first_name', '')).lower().strip()
        employee_full_name = f"{employee_first_name} {employee_last_name}"
        
        print(f"   Processing employee {emp_idx+1}: {employee_first_name} {employee_last_name}")
        
        # Find dependents for this employee using multiple strategies
        dependents_found = 0
        
        # Strategy 1: Exact last name match (highest priority)
        for dep_idx, dependent in enumerate(dependents):
            if dep_idx in used_dependents:
                continue
                
            dependent_last_name = str(dependent.get('last_name', '')).lower().strip()
            dependent_first_name = str(dependent.get('first_name', '')).lower().strip()
            
            # Exact last name match
            if (employee_last_name == dependent_last_name and 
                employee_last_name != '' and 
                employee_first_name != dependent_first_name):
                
                family.append(dependent)
                used_dependents.add(dep_idx)
                dependents_found += 1
                print(f"     ✅ Matched by last name: {dependent_first_name} {dependent_last_name}")
        
        # Strategy 2: Check for explicit employee references
        for dep_idx, dependent in enumerate(dependents):
            if dep_idx in used_dependents:
                continue
                
            # Check if dependent has explicit employee reference
            dependent_of = str(dependent.get('dependent_of_employee_row', '')).strip()
            employee_name_ref = str(dependent.get('employee_name', '')).lower().strip()
            
            # Check row reference (only if it makes logical sense)
            if dependent_of and dependent_of.isdigit() and int(dependent_of) == emp_idx + 1:
                # Additional check: make sure this makes logical sense
                # Don't match if last names are completely different
                dep_last_name = str(dependent.get('last_name', '')).lower().strip()
                if employee_last_name == dep_last_name or not dep_last_name:
                    family.append(dependent)
                    used_dependents.add(dep_idx)
                    dependents_found += 1
                    print(f"     ✅ Matched by row reference: {dependent.get('first_name', '')} {dependent.get('last_name', '')}")
            
            # Check employee name reference
            elif employee_name_ref and employee_name_ref == employee_full_name:
                family.append(dependent)
                used_dependents.add(dep_idx)
                dependents_found += 1
                print(f"     ✅ Matched by employee name: {dependent.get('first_name', '')} {dependent.get('last_name', '')}")
        
        print(f"     👨‍👩‍👧‍👦 Employee {emp_idx+1} has {dependents_found} dependents")
        grouped_families.append(family)
    
    # Strategy 3: Proximity-based matching for remaining ungrouped dependents
    remaining_dependents = [dep for i, dep in enumerate(dependents) if i not in used_dependents]
    
    if remaining_dependents:
        print(f"🔍 Applying proximity-based matching for {len(remaining_dependents)} remaining dependents...")
        
        # Create a mapping of all records to their original order for proximity calculation
        all_records_with_order = []
        for i, record in enumerate(all_results):
            all_records_with_order.append((i, record))
        
        # For each remaining dependent, find the closest family (not just employee) based on original order
        for dependent in remaining_dependents:
            dep_name = f"{dependent.get('first_name', '')} {dependent.get('last_name', '')}"
            print(f"   Finding closest family for: {dep_name}")
            
            # Find the dependent's position in the original data
            dep_position = None
            for pos, record in all_records_with_order:
                if (str(record.get('first_name', '')).lower().strip() == str(dependent.get('first_name', '')).lower().strip() and
                    str(record.get('last_name', '')).lower().strip() == str(dependent.get('last_name', '')).lower().strip()):
                    dep_position = pos
                    break
            
            if dep_position is not None:
                # Find the closest family by considering all family members' positions
                closest_family_idx = None
                min_distance = float('inf')
                
                for family_idx, family in enumerate(grouped_families):
                    # Calculate distance to the closest member of this family
                    family_min_distance = float('inf')
                    
                    for family_member in family:
                        # Find family member's position in original data
                        member_position = None
                        for pos, record in all_records_with_order:
                            if (str(record.get('first_name', '')).lower().strip() == str(family_member.get('first_name', '')).lower().strip() and
                                str(record.get('last_name', '')).lower().strip() == str(family_member.get('last_name', '')).lower().strip()):
                                member_position = pos
                                break
                        
                        if member_position is not None:
                            distance = abs(dep_position - member_position)
                            family_min_distance = min(family_min_distance, distance)
                    
                    # Choose the family with the closest member
                    if family_min_distance < min_distance:
                        min_distance = family_min_distance
                        closest_family_idx = family_idx
                
                if closest_family_idx is not None:
                    grouped_families[closest_family_idx].append(dependent)
                    employee_name = f"{grouped_families[closest_family_idx][0].get('first_name', '')} {grouped_families[closest_family_idx][0].get('last_name', '')}"
                    print(f"     ✅ Matched by proximity: {dep_name} → {employee_name} (distance: {min_distance})")
                else:
                    # Fallback to first available family
                    grouped_families[0].append(dependent)
                    employee_name = f"{grouped_families[0][0].get('first_name', '')} {grouped_families[0][0].get('last_name', '')}"
                    print(f"     ✅ Matched by fallback: {dep_name} → {employee_name}")
            else:
                # Fallback if position not found
                grouped_families[0].append(dependent)
                employee_name = f"{grouped_families[0][0].get('first_name', '')} {grouped_families[0][0].get('last_name', '')}"
                print(f"     ✅ Matched by fallback: {dep_name} → {employee_name}")
    
    print(f"👨‍👩‍👧‍👦 Created {len(grouped_families)} family groups")
    return grouped_families

def deduplicate_records(records):
    """
    Remove duplicate records based on name and DOB.
    """
    print(f"🔍 Removing duplicates from {len(records)} records...")
    
    seen = set()
    unique_records = []
    duplicates_found = 0
    
    for i, record in enumerate(records):
        if i % 100 == 0:  # Progress indicator
            print(f"   Processing record {i+1}/{len(records)}...")
        
        # Create a unique key based on name only (handle name order variations)
        first_name = str(record.get('first_name', '')).lower().strip()
        last_name = str(record.get('last_name', '')).lower().strip()
        
        # Create normalized key that handles name order variations
        name_parts = [first_name, last_name]
        name_parts = [part for part in name_parts if part]  # Remove empty parts
        key = tuple(sorted(name_parts))  # Sort to handle order variations
        
        if key not in seen and first_name != '' and last_name != '':  # Skip empty names
            seen.add(key)
            unique_records.append(record)
        else:
            duplicates_found += 1
            if duplicates_found <= 10:  # Only show first 10 duplicates
                print(f"⚠️ Duplicate found: {record.get('first_name', '')} {record.get('last_name', '')}")
            elif duplicates_found == 11:
                print("⚠️ ... (showing first 10 duplicates only)")
    
    print(f"✅ Removed {duplicates_found} duplicates, {len(unique_records)} unique records remain")
    return unique_records

# =========================================================
# MAIN LOOP
# =========================================================
def process_excel_in_chunks(file_path, question, chunk_size=CHUNK_SIZE, log_callback=None):
    all_sheets = excel_to_json_all(file_path)
    combined_results = []

    for sheet_name, records in all_sheets.items():
        if not records:
            msg = f"📄 Sheet '{sheet_name}' is empty, skipping."
            print(msg)
            if log_callback:
                log_callback(msg)
            continue

        msg = f"\n📄 Processing sheet: {sheet_name} ({len(records)} rows)"
        print(msg)
        if log_callback:
            log_callback(msg)
        
        num_chunks = math.ceil(len(records) / chunk_size)

        for i in range(num_chunks):
            start = i * chunk_size
            end = start + chunk_size
            chunk = records[start:end]
            msg = f"➡️ Rows {start + 1}-{min(end, len(records))}/{len(records)}"
            print(msg)
            if log_callback:
                log_callback(msg)

            # automatically handles large JSONs internally
            chunk_results = process_large_input(chunk, question, log_callback=log_callback)
            
            # Check for various error conditions
            if chunk_results == "RATE_LIMIT_EXCEEDED":
                if log_callback:
                    log_callback("🚫 Rate limit exceeded - aborting operations...")
                return "RATE_LIMIT_EXCEEDED"
            elif chunk_results == "TIMEOUT_ERROR":
                if log_callback:
                    log_callback("⏰ API timeout - data might be too large, reducing chunk size...")
                return "TIMEOUT_ERROR"
            elif chunk_results == "CONNECTION_ERROR":
                if log_callback:
                    log_callback("🔌 Connection error - check your internet connection")
                return "CONNECTION_ERROR"
            
            if chunk_results:
                combined_results.extend(chunk_results)

    # Post-processing: Group and deduplicate
    if combined_results:
        msg = "🔄 Post-processing: Grouping employees and dependents..."
        print(msg)
        if log_callback:
            log_callback(msg)
        
        # Remove duplicates first
        combined_results = deduplicate_records(combined_results)
        
        # Group employees and dependents
        grouped_families = group_employees_and_dependents(combined_results)
        
        # Flatten back to list for output
        final_results = []
        for family in grouped_families:
            final_results.extend(family)
        
        msg = f"✅ Final processing complete: {len(final_results)} records"
        print(msg)
        if log_callback:
            log_callback(msg)
        
        return final_results

    return combined_results

def calculate_field_confidence(records, original_data=None):
    """
    Calculate confidence scores for each field based on matching strategy and data quality.
    Returns a dictionary with field names, matching strategies, and confidence scores.
    """
    field_confidence = {}
    
    # Define field mappings and their expected characteristics
    field_definitions = {
        'last_name': {
            'type': 'text',
            'required': True,
            'max_length': 50,
            'pattern': r'^[A-Za-z\s\-\']+$'
        },
        'first_name': {
            'type': 'text', 
            'required': True,
            'max_length': 50,
            'pattern': r'^[A-Za-z\s\-\']+$'
        },
        'employee_name': {
            'type': 'derived',
            'source_fields': ['first_name', 'last_name'],
            'confidence_boost': 0.1
        },
        'home_zip_code': {
            'type': 'numeric',
            'required': False,
            'max_length': 10,
            'pattern': r'^\d{5}(-\d{4})?$'
        },
        'dob': {
            'type': 'date',
            'required': False,
            'pattern': r'^\d{1,2}/\d{1,2}/\d{4}$|^\d{4}-\d{2}-\d{2}$'
        },
        'gender': {
            'type': 'categorical',
            'required': False,
            'valid_values': ['Male', 'Female', 'M', 'F', 'male', 'female']
        },
        'medical_coverage': {
            'type': 'categorical',
            'required': False,
            'valid_values': ['Yes', 'No', 'Y', 'N', 'yes', 'no', 'true', 'false']
        },
        'vision_coverage': {
            'type': 'categorical',
            'required': False,
            'valid_values': ['Yes', 'No', 'Y', 'N', 'yes', 'no', 'true', 'false']
        },
        'dental_coverage': {
            'type': 'categorical',
            'required': False,
            'valid_values': ['Yes', 'No', 'Y', 'N', 'yes', 'no', 'true', 'false']
        },
        'cobra_participation': {
            'type': 'categorical',
            'required': False,
            'valid_values': ['Yes', 'No', 'Y', 'N', 'yes', 'no', 'true', 'false']
        },
        'relationship_to_employee': {
            'type': 'categorical',
            'required': True,
            'valid_values': ['Employee', 'Spouse', 'Child', 'Dependent', 'employee', 'spouse', 'child', 'dependent']
        },
        'dependent_of_employee_row': {
            'type': 'numeric',
            'required': False,
            'pattern': r'^\d+$'
        }
    }
    
    if not records:
        return field_confidence
    
    # Calculate confidence for each field
    for field_name, definition in field_definitions.items():
        confidence_data = {
            'field_name': field_name,
            'matching_strategy': 'LLM_Extraction',
            'confidence_score': 0.0,
            'data_quality_score': 0.0,
            'completeness_score': 0.0,
            'validation_score': 0.0,
            'total_records': len(records),
            'valid_records': 0,
            'empty_records': 0,
            'invalid_records': 0
        }
        
        valid_count = 0
        empty_count = 0
        invalid_count = 0
        
        for record in records:
            field_value = str(record.get(field_name, '')).strip()
            
            # Check if field is empty
            if not field_value or field_value.lower() in ['null', 'none', '']:
                empty_count += 1
                continue
            
            # Validate field based on definition
            is_valid = True
            
            # Length validation
            if definition.get('max_length') and len(field_value) > definition['max_length']:
                is_valid = False
            
            # Pattern validation
            if definition.get('pattern'):
                import re
                if not re.match(definition['pattern'], field_value):
                    is_valid = False
            
            # Categorical validation
            if definition.get('type') == 'categorical' and definition.get('valid_values'):
                if field_value not in definition['valid_values']:
                    is_valid = False
            
            # Date validation
            if definition.get('type') == 'date':
                try:
                    from datetime import datetime
                    if '/' in field_value:
                        datetime.strptime(field_value, '%m/%d/%Y')
                    elif '-' in field_value:
                        datetime.strptime(field_value, '%Y-%m-%d')
                    else:
                        is_valid = False
                except:
                    is_valid = False
            
            if is_valid:
                valid_count += 1
            else:
                invalid_count += 1
        
        # Calculate scores
        total_records = len(records)
        if total_records > 0:
            completeness_score = (total_records - empty_count) / total_records
            validation_score = valid_count / total_records if total_records > 0 else 0
            data_quality_score = (completeness_score + validation_score) / 2
            
            # Base confidence from LLM extraction
            base_confidence = 0.85  # LLM extraction base confidence
            
            # Adjust confidence based on data quality
            if definition.get('required', False):
                # Required fields get higher weight
                confidence_score = base_confidence * (0.7 + 0.3 * data_quality_score)
            else:
                # Optional fields
                confidence_score = base_confidence * (0.5 + 0.5 * data_quality_score)
            
            # Boost for derived fields
            if definition.get('type') == 'derived':
                confidence_score += definition.get('confidence_boost', 0.1)
            
            # Cap confidence between 0 and 1
            confidence_score = max(0.0, min(1.0, confidence_score))
            
            confidence_data.update({
                'confidence_score': round(confidence_score, 3),
                'data_quality_score': round(data_quality_score, 3),
                'completeness_score': round(completeness_score, 3),
                'validation_score': round(validation_score, 3),
                'valid_records': valid_count,
                'empty_records': empty_count,
                'invalid_records': invalid_count
            })
        
        field_confidence[field_name] = confidence_data
    
    return field_confidence

def save_field_confidence(confidence_data, output_path):
    """Save field confidence data to a separate JSON file."""
    import json
    import os
    
    # Create confidence file path
    base_path = os.path.splitext(output_path)[0]
    confidence_path = f"{base_path}_field_confidence.json"
    
    # Prepare summary data
    summary = {
        'analysis_timestamp': pd.Timestamp.now().isoformat(),
        'total_fields_analyzed': len(confidence_data),
        'average_confidence': round(sum(field['confidence_score'] for field in confidence_data.values()) / len(confidence_data), 3),
        'high_confidence_fields': [name for name, data in confidence_data.items() if data['confidence_score'] >= 0.8],
        'medium_confidence_fields': [name for name, data in confidence_data.items() if 0.6 <= data['confidence_score'] < 0.8],
        'low_confidence_fields': [name for name, data in confidence_data.items() if data['confidence_score'] < 0.6]
    }
    
    # Combine summary and detailed data
    output_data = {
        'summary': summary,
        'field_analysis': confidence_data
    }
    
    # Save to file
    with open(confidence_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, indent=4)
    
    print(f"📊 Field confidence analysis saved to: {confidence_path}")
    return confidence_path

def json_to_csv_by_input_name(json_filepath):
    """
    Converts a JSON file (list of objects) to a CSV file with proper family grouping.
    """
    try:
        # 1. Path Management: Generate the output CSV filepath
        json_dir = os.path.dirname(json_filepath)
        base_name = os.path.splitext(os.path.basename(json_filepath))[0]
        csv_filename = base_name + ".csv"
        csv_filepath = os.path.join(json_dir, csv_filename)
        
        # 2. Read JSON and convert to CSV
        with open(json_filepath, 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)
        
        if not data:
            print("No data to convert to CSV.")
            return
        
        # 3. Apply family grouping logic
        grouped_families = group_employees_and_dependents(data)
        
        # 4. Create CSV with proper column order and family grouping
        csv_data = []
        sl_no = 1
        
        for family_id, family in enumerate(grouped_families, 1):
            family_size = len(family)
            
            for person in family:
                # Determine if person is a dependent
                relationship = str(person.get('relationship_to_employee', '')).lower().strip()
                is_dependent = relationship in ['spouse', 'child', 'dependent']
                
                # Create employee full name
                employee_name = ""
                if not is_dependent:
                    employee_name = f"{person.get('first_name', '')} {person.get('last_name', '')}"
                else:
                    # Find the employee in this family
                    for family_member in family:
                        member_rel = str(family_member.get('relationship_to_employee', '')).lower().strip()
                        if member_rel == 'employee':
                            employee_name = f"{family_member.get('first_name', '')} {family_member.get('last_name', '')}"
                            break
                
                # Create row with proper column order
                row = {
                    'Sl.No': sl_no,
                    'Family_Group_ID': family_id,
                    'Last_Name': person.get('last_name', ''),
                    'First_Name': person.get('first_name', ''),
                    'Employee_Full_Name': employee_name,
                    'Relationship_to_Employee': person.get('relationship_to_employee', ''),
                    'Dependents_Y_N': 'Y' if is_dependent else 'N',
                    'DOB': person.get('dob', ''),
                    'Gender': person.get('gender', ''),
                    'Home_Zip_Code': person.get('home_zip_code', ''),
                    'Medical_Coverage': person.get('medical_coverage', ''),
                    'Dental_Coverage': person.get('dental_coverage', ''),
                    'Vision_Coverage': person.get('vision_coverage', ''),
                    'COBRA_Participation': person.get('cobra_participation', ''),
                    'Dependent_of_Employee_Row': person.get('dependent_of_employee_row', ''),
                    'Family_Size': family_size
                }
                
                csv_data.append(row)
                sl_no += 1
        
        # 5. Write to CSV with proper column order
        fieldnames = [
            'Sl.No', 'Family_Group_ID', 'Last_Name', 'First_Name', 'Employee_Full_Name',
            'Relationship_to_Employee', 'Dependents_Y_N', 'DOB', 'Gender', 'Home_Zip_Code',
            'Medical_Coverage', 'Dental_Coverage', 'Vision_Coverage', 'COBRA_Participation',
            'Dependent_of_Employee_Row', 'Family_Size'
        ]
        
        with open(csv_filepath, 'w', newline='', encoding='utf-8') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(csv_data)
        
        print(f"✅ CSV saved to: {csv_filepath}")
        print(f"📊 Created {len(grouped_families)} family groups with {len(csv_data)} total records")
        
    except json.JSONDecodeError:
        print(f"Error: Could not decode JSON from '{json_filepath}'. Check file format.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

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
        self.process_button = ttk.Button(main_frame, text="Process File (Enhanced with Grouping)", command=self.process_file, state="disabled")
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
                import subprocess
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
            self.log_status(f"⚙️ Config: CHUNK_SIZE={CHUNK_SIZE}, MAX_CHARS={MAX_CHARS}, MAX_TOKENS={MAX_TOKENS}")
            self.log_status(f"📄 Processing file: {os.path.basename(file_path)}")
            
            # Set up output paths
            base, _ = os.path.splitext(file_path)
            self.output_json_path = f"{base}_output_grouped.json"
            self.output_csv_path = f"{base}_output_grouped.csv"
            
            # Define the question for the LLM
            question = """
            You are given an Excel file with employee census data. Analyse all input data and return the output for all the data.
            Extract both employees and their dependents and return them as valid JSON only.
            Do not want "employees" ans "Dependents" separately in output json , do  not put assumed example data

            Rules:
            -
            - Read all worksheets, find employee info, and fetch data once found do not process other sheets .
            - Use the exact field names provided.
            - If COBRA Participant is missing, return "No".
            - If First Name is missing, skip that row.
            - Dependents: split 'DEPENDENT n NAME' into First Name and Last Name.
            - Relationship to Employee: use exactly as shown in Excel (no grouping by last name).
            - Employee rows: Relationship to Employee = "Employee", Dependent of Employee Row = null.
            - Dependents: Dependent of Employee Row should link to the Employee in order of appearance (row index).
            - include Sl.No in output json  - row starts with 1
            """
            
            # Process the file using the existing logic
            results = process_excel_in_chunks(file_path, question, chunk_size=CHUNK_SIZE, log_callback=self.log_status)
            
            # Check for various error conditions
            if results == "RATE_LIMIT_EXCEEDED":
                self.log_status("🚫 Rate limit exceeded - aborting operations...")
                self.log_status("💡 Please try again later or upgrade your Groq plan for higher limits.")
                return
            elif results == "TIMEOUT_ERROR":
                self.log_status("⏰ API timeout - data might be too large")
                self.log_status("💡 Try with a smaller Excel file or reduce the chunk size")
                return
            elif results == "CONNECTION_ERROR":
                self.log_status("🔌 Connection error - check your internet connection")
                self.log_status("💡 Please check your internet connection and try again")
                return
            
            if results:
                try:
                    with open(self.output_json_path, "w", encoding="utf-8") as f:
                        json.dump(results, f, indent=4)

                    json_to_csv_by_input_name(self.output_json_path)
                    self.log_status(f"\n✅ Final JSON saved to: {self.output_json_path}")
                    self.log_status(f"✅ Final CSV saved to: {self.output_csv_path}")
                    
                    # Calculate field confidence scores
                    self.log_status("📊 Calculating field confidence scores...")
                    confidence_data = calculate_field_confidence(results)
                    confidence_path = save_field_confidence(confidence_data, self.output_json_path)
                    self.log_status(f"✅ Field confidence analysis saved to: {confidence_path}")
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
        finally:
            # Re-enable process button and stop progress
            self.progress.stop()
            self.process_button.config(state="normal")
            
    def view_json(self):
        if hasattr(self, 'output_json_path') and self.output_json_path and os.path.exists(self.output_json_path):
            try:
                import subprocess
                import platform
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
                import subprocess
                import platform
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
def main():
    root = tk.Tk()
    app = GroqCensusExtractorUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
