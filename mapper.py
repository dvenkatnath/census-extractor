import os, json, csv, io, string
from groq import Groq
from dotenv import load_dotenv
from learning_system import learning_system

# Load environment variables from .env file
load_dotenv()

api_key = os.getenv("GROQ_API_KEY")
if api_key:
    print(f"🔑 API Key loaded from .env file: {api_key[:10]}...")
else:
    print("❌ No API key found in .env file!")

client = Groq(api_key=api_key)

def _convert_column_letters_to_names(mapping: dict, thin_csv: str) -> dict:
    """Convert column letters (A, B, C) to actual column names and validate content-based mappings"""
    print("🔄 Starting post-processing...")
    
    # Parse the CSV to get column names and data
    lines = thin_csv.strip().split('\n')
    if not lines:
        print("❌ No CSV data found")
        return mapping
    
    # Get header row (first line after __sheet__)
    header_line = None
    for line in lines:
        if line.startswith('__sheet__,'):
            header_line = line
            break
    
    if not header_line:
        print("❌ No header row found")
        return mapping
    
    # Extract column names (skip the first __sheet__ column)
    column_names = header_line.split(',')[1:]  # Skip __sheet__
    print(f"🔍 Found columns: {column_names}")
    
    # Create mapping from letter to column name
    letter_to_name = {}
    for i, col_name in enumerate(column_names):
        if i < 26:  # Only handle A-Z
            letter = string.ascii_uppercase[i]
            letter_to_name[letter] = col_name.strip()
    
    print(f"🔍 Letter mapping: {letter_to_name}")
    
    # Get sample data rows for content validation
    data_rows = []
    for line in lines[1:6]:  # First 5 data rows
        if line and not line.startswith('__sheet__'):
            data_rows.append(line.split(','))
    
    print(f"🔍 Sample data rows: {len(data_rows)}")
    
    # Convert the mapping and validate content
    converted_mapping = {}
    for field, refs in mapping.items():
        print(f"🔄 Processing field: {field} with refs: {refs}")
        converted_refs = []
        for ref in refs:
            if ',' in ref:
                sheet_name, col_ref = ref.split(',', 1)
                col_ref = col_ref.strip()
                
                # If it's a single letter, convert it
                if len(col_ref) == 1 and col_ref in letter_to_name:
                    converted_ref = f"{sheet_name},{letter_to_name[col_ref]}"
                    print(f"🔄 Converted '{ref}' to '{converted_ref}'")
                    converted_refs.append(converted_ref)
                else:
                    print(f"🔄 Keeping '{ref}' as is")
                    converted_refs.append(ref)
            else:
                converted_refs.append(ref)
        
        # Content-based validation for Relationship To employee
        if field == "Relationship To employee":
            print(f"🎯 Validating Relationship To employee mapping...")
            corrected_refs = _validate_relationship_mapping(converted_refs, column_names, data_rows)
            if corrected_refs != converted_refs:
                print(f"🎯 Content-based correction: {converted_refs} → {corrected_refs}")
                converted_refs = corrected_refs
            else:
                print(f"🎯 No correction needed for Relationship To employee")
        
        converted_mapping[field] = converted_refs
    
    print(f"🔄 Final converted mapping: {converted_mapping}")
    return converted_mapping

def _validate_relationship_mapping(refs: list, column_names: list, data_rows: list) -> list:
    """Validate Relationship To employee mapping based on content - trust LLM mapping, just verify"""
    if not refs or not data_rows:
        return refs
    
    # First, validate that the mapped column actually contains relationship values
    # Parse the current mapping to get the sheet and column
    mapped_sheet = None
    mapped_col = None
    
    for ref in refs:
        if ',' in ref:
            parts = ref.split(',', 1)
            mapped_sheet = parts[0].strip()
            mapped_col = parts[1].strip()
            break
    
    # If we have a mapped column, verify it contains relationship values
    if mapped_col:
        try:
            col_index = column_names.index(mapped_col)
            # Check if this column actually contains relationship values
            relationship_score = 0
            job_title_score = 0
            
            for row in data_rows:
                if col_index < len(row):
                    value = str(row[col_index]).strip().upper()
                    # Check for relationship keywords
                    if value in ['SPOUSE', 'CHILD', 'EMPLOYEE', 'SELF', 'DEPENDENT', 'WIFE', 'HUSBAND', 'SON', 'DAUGHTER']:
                        relationship_score += 3
                    elif any(keyword in value for keyword in ['SPOUSE', 'CHILD', 'DEPENDENT']):
                        relationship_score += 2
                    # Check for job title keywords (which should NOT be in relationship column)
                    elif any(keyword in value for keyword in ['PATIENT CARE', 'MANAGER', 'ASSISTANT', 'ACCOUNTANT', 'NURSE', 'DOCTOR', 'DIRECTOR']):
                        job_title_score += 1
            
            # If the mapped column has more job title keywords than relationship keywords, 
            # it might be wrong - but trust LLM if it's close
            if job_title_score > relationship_score * 2:
                print(f"⚠️ Warning: Mapped column '{mapped_col}' appears to contain job titles, not relationships")
                # Don't override - let user correct if needed
        except (ValueError, IndexError):
            # Column not found or index error - keep original mapping
            pass
    
    # Trust the LLM mapping - return it as-is
    # The LLM has already analyzed the content, so trust its judgment
    return refs

def build_mapping(thin_csv: str, canonical: list[str], file_name: str = "unknown") -> dict[str, list[str]]:
    """
    Returns dict  canonical_field -> list["sheet,col_name", …]
    """
    print("🔗 Connected to Llama 3.3 70B - Versatile")
    print("📤 Sending first 5 rows of sample data...")
    print("📊 Data being sent to Groq API:")
    print("=" * 50)
    print(thin_csv)
    print("=" * 50)
    
    # Extract column names from CSV for learning context
    lines = thin_csv.strip().split('\n')
    if lines:
        column_names = lines[0].split(',')[1:]  # Skip __sheet__ column
    else:
        column_names = []
    
    # Get learning context
    learning_context = learning_system.get_learning_context(column_names, thin_csv)
    
    # Debug: Print learning context
    if learning_context:
        print("🧠 Learning context being applied:")
        print(learning_context)
        print("=" * 50)
    else:
        print("🧠 No learning context available")
        print("=" * 50)
    
    prompt = f"""You are a census-data schema expert. Analyze the Excel data below and map columns to standard census fields.

    {learning_context}

    Data (first 5 rows of each sheet):
    {thin_csv}

    🚨 CRITICAL INSTRUCTIONS:
    1. **ALWAYS use "SheetName,ColumnName" format** - NEVER use column letters like A, B, C
    2. **Look at the ACTUAL VALUES in each column** to determine the mapping - content determines mapping, not headers!
    3. **For "Relationship To employee"**: Look for columns containing values like "Spouse", "Child", "Employee", "Self", "PATIENT CARE ASSISTANT" - these indicate family relationships or job titles that show relationship
    4. **For "Medical Coverage"**: Look for columns with values like "Employee", "E", "Employee + Spouse", "F", "ES" - these indicate coverage levels
    5. **Content determines mapping, not column headers** - if a column has "Job Title" as header but contains "Spouse", "Child" values, it should map to "Relationship To employee"
    6. **Return ONLY a valid JSON object** - NO explanations, NO text, NO markdown, just pure JSON
    7. If no suitable column exists, use "UNKNOWN"

    🚨 CRITICAL NAME HANDLING RULES:
    - **FULL NAME COLUMNS**: If a column contains full names (e.g., "John Smith", "Mary Jane Doe", "Smith, John"):
      * Map it to BOTH "First Name" AND "Last Name" columns
      * Also map it to "Employee Name" if available
      * The system will automatically split the full name into first and last name
    - **SEPARATE NAME COLUMNS**: If there are separate "First Name" and "Last Name" columns, map them directly
    - **MIXED FORMATS**: Handle both formats:
      * If "Employee Name" has full names like "John Smith" → map to both "First Name" and "Last Name"
      * If "Name" column has full names → map to both "First Name" and "Last Name"
      * If separate "First" and "Last" columns exist → use them directly
    - **Name splitting logic**: Last word = Last Name, All other words = First Name
      * Example: "Mary Jane Doe" → First Name: "Mary Jane", Last Name: "Doe"
      * Example: "John Smith" → First Name: "John", Last Name: "Smith"
      * Example: "Smith, John" → First Name: "John", Last Name: "Smith"

    Key patterns to recognize:
    - Relationship to employee: Look for columns with values like "Spouse", "Child", "Dependent", or job titles that indicate relationships
    - Medical Coverage: Look for columns with values like "Employee", "E", "Employee only", "Employee + Spouse", "F", "ES", etc.
    - First Name: Look for columns with personal names (first part of full name if full name column)
    - Last Name: Look for columns with surnames (last part of full name if full name column)
    - Employee Name: Look for columns with full names (can be mapped to both First Name and Last Name)
    - DOB: Look for date patterns (MM/DD/YYYY, etc.)
    - Gender: Look for "M", "F", "Male", "Female", etc.

    Official census fields:
    {json.dumps(canonical)}

    Required JSON format:
    {{"Field Name": ["SheetName,ColumnName"], "Another Field": ["Sheet1,Column1", "Sheet2,Column2"], ...}}

    Example:
    {{"First Name": ["Census,First"], "Last Name": ["Census,Employee  Name"], "DOB": ["Census,DOB"], "Relationship To employee": ["Census,Job Title"]}}

    🚨 CRITICAL: Use ColumnName (like "First", "Employee Name", "Job Title") NOT column letters (like "A", "B"). 
    🚨 CRITICAL: Look at the actual data values to make your decision - content determines mapping, not headers!
    🚨 CRITICAL: If you see "Spouse", "Child" in any column, that column maps to "Relationship To employee"!
    🚨 CRITICAL: Return ONLY JSON - NO explanations, NO text, NO markdown!

    RESPOND WITH ONLY THIS JSON FORMAT:
    {{"First Name": ["Census,First"], "Last Name": ["Census,Employee  Name"], "DOB": ["Census,DOB"], "Gender": ["Census,Gender"], "Relationship To employee": ["Census,Role"], "Medical Coverage": ["Census,Coverage Level"], "Medical Plan Name": ["Census,Healthcare"]}}"""
    
    print("🤖 Sending request to Groq API...")
    reply = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.0, max_tokens=1500
    )
    
    print("✅ Received response from Groq API")
    print("📋 Parsing JSON response...")
    
    content = reply.choices[0].message.content
    print(f"🔍 Raw API response length: {len(content) if content else 0}")
    print(f"🔍 Raw API response: {repr(content)}")
    
    if not content or content.strip() == "":
        print("⚠️  Warning: Empty response from Groq API")
        return {}
    
    # Extract JSON from response (handle explanatory text)
    print("🧹 Extracting JSON from response...")
    
    # Try to extract JSON from the response
    if '```json' in content:
        # Extract content between ```json and ```
        start = content.find('```json') + 7
        end = content.find('```', start)
        if end != -1:
            content = content[start:end].strip()
            print(f"🧹 Extracted from ```json: '{content}'")
    elif '```' in content:
        # Extract content between ``` and ```
        start = content.find('```') + 3
        end = content.find('```', start)
        if end != -1:
            content = content[start:end].strip()
            print(f"🧹 Extracted from ```: '{content}'")
    
    # Look for JSON object in the content
    if '{' in content and '}' in content:
        start = content.find('{')
        end = content.rfind('}') + 1
        content = content[start:end]
        print(f"🧹 Extracted JSON object: '{content}'")
    
    try:
        result = json.loads(content)
        print("✅ Successfully parsed JSON response")
        print(f"📋 Parsed result: {result}")
        
        # Post-process to convert column letters to column names and validate content
        result = _convert_column_letters_to_names(result, thin_csv)
        print(f"🔄 Post-processed result: {result}")
        
        return result
    except json.JSONDecodeError as e:
        print(f"❌ Warning: Failed to parse JSON from Groq API: {e}")
        print(f"📄 Raw response: {content}")
        print(f"📄 Response type: {type(content)}")
        return {}

def store_successful_mapping(original_mapping: dict, corrected_mapping: dict, 
                           column_names: list, sample_data: str, file_name: str = "unknown"):
    """Store a successful mapping for learning"""
    learning_system.store_successful_mapping(
        original_mapping, corrected_mapping, column_names, sample_data, file_name
    )