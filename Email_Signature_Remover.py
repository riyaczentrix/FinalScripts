#pip install convosense_utilities
#pip install email_reply_parser
#pip install openpyxl


import pandas as pd
import re
import numpy as np
import sys
import os
import nltk

# --- NLTK Data Download (Run these lines once if you haven't downloaded them) ---
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    print("NLTK 'punkt' tokenizer not found, downloading...")
    nltk.download('punkt', quiet=True)
try:
    nltk.data.find('tokenizers/punkt_tab')
except LookupError:
    print("NLTK 'punkt_tab' tokenizer not found, downloading...")
    nltk.download('punkt_tab', quiet=True)
try:
    nltk.data.find('corpora/wordnet')
except LookupError:
    print("NLTK 'wordnet' corpus not found, downloading...")
    nltk.download('wordnet', quiet=True)


# --- Import actual email_signature_remover or email_reply_parser ---
# This block attempts to import convosense_utilities first, then falls back to email_reply_parser,
# and finally provides a basic mock if neither is available.
try:
    # Attempt to import convosense_utilities (if available in the environment)
    from convosense_utilities.email_signature_remover import remove_sign
    print("Using 'convosense_utilities.email_signature_remover.remove_sign'.")
except ImportError:
    try:
        # Fallback to email_reply_parser if convosense_utilities is not found
        from email_reply_parser import EmailReplyParser
        print("Warning: 'convosense_utilities' not found. Using 'email_reply_parser'.")
        class EmailSignatureRemoverProxy:
            def remove_sign(self, text):
                if not isinstance(text, str):
                    return ""
                # email_reply_parser.EmailReplyParser.parse_reply is the core function
                return EmailReplyParser.parse_reply(text)
        remove_sign = EmailSignatureRemoverProxy().remove_sign # Assign the method directly
    except ImportError:
        # If neither library is available, raise an error as per user's strict requirement
        print("Error: Neither 'convosense_utilities' nor 'email_reply_parser' found. "
              "Cannot proceed without a suitable email signature removal library. "
              "Please ensure one of these libraries is installed (e.g., pip install email_reply_parser convosense_utilities).", file=sys.stderr)
        sys.exit(1)


# --- Configuration ---
FILE_NAME = 'ticket_details_5000.xlsx' # Your input XLSX file
PROBLEM_COLUMN = 'problem_reported' # Name of the problem reported column
MAIL_BODY_COLUMN = 'DecodedBody' # Original column name for mail body
CLEANED_OUTPUT_FILE = 'temp_file.csv' # Output file for cleaned data


# --- Pre-compile regex patterns for performance ---

# Specific multi-line pattern to be completely removed from DecodedBody (User Request 3)
# Note: The 'b' prefix in the user's example indicates a byte string, but Python strings
# will be matched against this. Newlines and spaces are important for this pattern.
# This pattern is retained for explicit removal as it's a known problematic block.
SPECIFIC_DECODED_BODY_PATTERN = r"""
Internal 
 
 
 
 
 Dear Ravi, 
 
 Kindly extend your support to share the Not Dialed reason for below list. Be noted these numbers were uploaded 4 times on 9th Feb’24 but was not dialed at all. If it was rescheduled, I need the id with which it rescheduled.
 
 
 
 
 
 
 Not Dialed 
 
 
 
 
 599075055 
 
 
 
 
 553575999 
 
 
 
 
 500700044 
 
 
 
 
 535216630 
 
 
 
 
 555234049 
 
 
 
 
 539999495 
 
 
 
 
 567582101 
 
 
 
 
 533855394 
 
 
 
 
 503882099 
 
 
 
 
 574290175 
 
 
 
 
 
 549344233 
 
 
 
 
 535054455 
 
 
 
 
 555594490 
 
 
 
 
 557908696 
 
 
 
 
 599798668 
 
 
 
 
 594999991 
 
 
 
 
 566212252 
 
 
 
 
 569744688 
 
 
 
 
 508021971 
 
 
 
 
 533783672 
 
 
 
 
 509374590 
 
 
 
 
 503393098 
 
 
 
 
 530660812 
 
 
 
 
 505944473 
 
 
 
 
 555593243 
 
 
 
 
 591794555 
 
 
 
 
 544501684 
 
 
 
 
 555302289 
 
 
 
 
 564637219 
 
 
 
 
 530805808 
 
"""
# Normalize whitespace in the pattern for robust matching
_SPECIFIC_BLOCK_NORMALIZED = re.sub(r'\s+', ' ', SPECIFIC_DECODED_BODY_PATTERN).strip().lower()

# Keywords to insert full stop and newlines around (User Request 2 & 9)
# Order matters for regex: put longer phrases first to ensure they are matched as a whole.
_NEWLINE_INSERTION_KEYWORDS = [
    r'Thanks\s*&\s*Regards',    # Specific "Thanks & Regards" with ampersand
    r'Thanks\s+and\s+Regards',  # "Thanks and Regards" with "and"
    r'Best\s*Regards',          # "Best Regards"
    r'Warm\s*Regards',          # "Warm Regards" - Added
    r'Regards,?',               # "Regards" or "Regards," - Adjusted to capture optional comma
    r'Thanks',                  # Individual "Thanks"
    r'Thank'                    # Individual "Thank"
]
# Create a combined pattern to capture the keyword itself (including optional trailing comma)
_COMBINED_KEYWORDS_PATTERN = r'\b(' + '|'.join(_NEWLINE_INSERTION_KEYWORDS) + r')\b'


# --- Cleaning Functions ---

def clean_problem_reported_entry(problem_text):
    """
    Cleans a single problem reported entry based on specified rules.
    Returns None if the entry should be removed (meaning the entire row is dropped),
    otherwise returns the cleaned text.

    Rules applied (User Requests 1, 5, 6, 7):
    - Remove if starts with 're', 'fwd:', 'fw' (case insensitive).
    - Remove if contains 'Mail Delivery Failure' or 'Mail delivery failed' (case insensitive).
    - Remove if contains 'Undeliverable' (case insensitive).
    - Remove if contains 'Resolved', 'Warning', 'Disaster', 'High' (case insensitive).
    """
    if pd.isna(problem_text) or not str(problem_text).strip():
        return None # Remove if NaN or empty/whitespace

    text = str(problem_text).strip()

    # Rule 1: Remove if starts with 're', 'fwd:', 'fw' (case insensitive)
    if re.match(r'^(re|fwd:|fw)\s*:', text, re.IGNORECASE):
        return None

    # Rule 5: Remove if contains 'Mail Delivery Failure' or 'Mail delivery failed' (case insensitive)
    if re.search(r'mail delivery failure|mail delivery failed', text, re.IGNORECASE):
        return None

    # Rule 6: Remove if contains 'Undeliverable' (case insensitive)
    if re.search(r'undeliverable', text, re.IGNORECASE):
        return None

    # Rule 7: Remove if contains 'Resolved', 'Warning', 'Disaster', 'High' (case insensitive)
    if re.search(r'resolved|warning|disaster|high', text, re.IGNORECASE):
        return None

    return text # Return original text if no removal rule matched

def clean_decoded_body(text):
    """
    Cleans the email body by applying several rules:
    - Removes specific large block of text (User Request 3).
    - Inserts full stop and newlines after specific signature phrases (User Request 2 & 9).
    - Applies email_signature_remover for general reply chain and disclaimer removal.
    - Normalizes whitespace.
    Returns an empty string if the body becomes empty or non-meaningful after cleaning.
    """
    if not isinstance(text, str):
        return ""

    # Make a copy of the original text for `replace_signature_break` to reference
    original_text_for_ref = text.strip()
    cleaned_text = original_text_for_ref # Start with the original text for initial processing

    # --- Phase 0: Specific block and hard-coded reply/auto-generated message removal ---
    # This is kept as a pre-filter for known problematic blocks that might cause issues
    # for email_signature_remover or are explicitly defined for removal.
    normalized_current_text_for_comparison = re.sub(r'\s+', ' ', cleaned_text).strip().lower()
    if _SPECIFIC_BLOCK_NORMALIZED in normalized_current_text_for_comparison:
        return "" # Return empty string if the specific pattern is found (leading to row removal)

    # Check for specific reply/auto-generated lines that should trigger row removal
    if re.search(r'reply above this line', cleaned_text, re.IGNORECASE) or \
       re.search(r'this message was created automatically by mail delivery software', cleaned_text, re.IGNORECASE) or \
       re.search(r'‚Äî-‚Äî-‚Äî-‚Äî Reply above this line', cleaned_text, re.IGNORECASE):
        return "" # Return empty string to trigger row removal

    # --- Phase 1: Initial Replacements for Newline Feature (based on user examples) ---
    # These are text replacements, not signature removals, and are applied before formatting.
    cleaned_text = re.sub(r'\bsnip\b', 'Thanks and regards', cleaned_text, flags=re.IGNORECASE)
    cleaned_text = re.sub(r'\bDisclaimer\b', 'Thanks and regards', cleaned_text, flags=re.IGNORECASE)

    # --- Phase 2: Insert full stop and newlines BEFORE and AFTER specific signature phrases ---
    # This section now manually reconstructs the string to ensure precise formatting.
    formatted_parts = []
    last_end_index = 0

    for match in re.finditer(_COMBINED_KEYWORDS_PATTERN, cleaned_text, flags=re.IGNORECASE):
        keyword_matched = match.group(1) # The actual keyword (e.g., "Thanks", "Regards,")
        start_index = match.start()
        end_index = match.end()

        # Add text before the current match
        formatted_parts.append(cleaned_text[last_end_index:start_index])

        # Determine if a full stop needs to be inserted
        # Look at the content *before* the match, ignoring leading/trailing whitespace/newlines.
        text_before_match_point = cleaned_text[last_end_index:start_index]
        
        last_content_char_before_match = ''
        for i in range(len(text_before_match_point) - 1, -1, -1):
            char = text_before_match_point[i]
            if char.strip(): # If it's not whitespace
                last_content_char_before_match = char
                break

        if last_content_char_before_match and last_content_char_before_match not in ['.', '!', '?', ';']:
            formatted_parts.append(".")
        
        # Always add a newline before the keyword to ensure it starts on a new line
        formatted_parts.append("\n")

        # Process 'Thank' to 'Thanks'
        processed_keyword = keyword_matched
        if re.fullmatch(r'Thank', keyword_matched, re.IGNORECASE):
            processed_keyword = "Thanks"

        # Handle special split for "Thanks & Regards" / "Thanks and Regards"
        if re.search(r'Thanks\s*(?:&|and)\s*Regards', processed_keyword, re.IGNORECASE):
            match_split = re.search(r'(Thanks)\s*(&|and)\s*(Regards)', processed_keyword, re.IGNORECASE)
            if match_split:
                thanks_part = match_split.group(1).strip()
                connector_regards_part = f"{match_split.group(2).strip()} {match_split.group(3).strip()}"
                formatted_parts.append(thanks_part)
                formatted_parts.append("\n") # Newline after "Thanks" part
                formatted_parts.append(connector_regards_part)
            else:
                formatted_parts.append(processed_keyword.strip())
        else:
            formatted_parts.append(processed_keyword.strip())
        
        # Always add another newline after the keyword (or split parts)
        formatted_parts.append("\n")

        last_end_index = end_index

    # Add any remaining text after the last match
    formatted_parts.append(cleaned_text[last_end_index:])
    cleaned_text = "".join(formatted_parts)

    # --- Phase 3: Apply email_signature_remover for comprehensive general cleaning ---
    # This handles removal of general reply chains and disclaimers.
    cleaned_text = remove_sign(cleaned_text)

    # --- Final Normalization (whitespace, newlines) ---
    # Normalize multiple newlines to at most two (representing a paragraph break)
    cleaned_text = re.sub(r'\n{3,}', '\n\n', cleaned_text)
    
    # Strip leading/trailing whitespace from each line, then from the whole text
    text_lines_processed = []
    for line in cleaned_text.split('\n'):
        stripped_line = line.strip()
        if stripped_line:
            text_lines_processed.append(stripped_line)
        elif line: # Keep empty lines if they were not just whitespace
            text_lines_processed.append('')
    
    cleaned_text = "\n".join(text_lines_processed)
    cleaned_text = cleaned_text.strip()

    # --- Final check for "meaningful" content (after all cleaning) ---
    # If the text is now empty or only contains whitespace, return an empty string.
    if not cleaned_text:
        return ""

    return cleaned_text


# --- Main execution ---
print(f"Loading data from '{FILE_NAME}' with 'latin1' encoding...")
try:
    df = pd.read_excel(FILE_NAME, engine='openpyxl') # Changed to read_excel and added engine
    df = df.fillna('') # Fill NaN values with empty strings immediately after loading
    print(f"Initial rows loaded: {len(df)}")
    print("\n--- Head of RAW DataFrame before cleaning ---")
    print(df.head())
except FileNotFoundError:
    print(f"Error: '{FILE_NAME}' not found. Please ensure the file is in the same directory as the script or provide the full path.", file=sys.stderr)
    sys.exit(1)
except Exception as e: # Catching general Exception for read_excel errors
    print(f"An error occurred while loading the XLSX file: {e}", file=sys.stderr)
    print("Please ensure 'openpyxl' is installed (pip install openpyxl).", file=sys.stderr)
    sys.exit(1)

# Handle BOM character in 'docket_no' if present (from previous debugging)
if 'ï»¿docket_no' in df.columns:
    df.rename(columns={'ï»¿docket_no': 'docket_no'}, inplace=True)
    print("Renamed column 'ï»¿docket_no' to 'docket_no'.")


# Ensure required columns exist
required_columns = ['docket_no', 'problem_reported', 'priority_name',
                    'disposition_name', 'sub_disposition_name', 'DecodedBody',
                    'mail_list_id', 'mail_id', 'ticket_id', 'assigned_to_dept_name', 'assigned_to_user_name']

missing_columns = [col for col in required_columns if col not in df.columns]

if missing_columns:
    print(f"Error: The following required columns were not found in '{FILE_NAME}': {missing_columns}")
    print(f"Available columns are: {df.columns.tolist()}")
    sys.exit(1)


# --- Apply Cleaning ---
print("\nApplying cleaning rules...")

# Apply cleaning for Problem Reported column (User Requests 1, 5, 6, 7)
initial_rows_count = len(df)
df['temp_problem_reported'] = df[PROBLEM_COLUMN].apply(clean_problem_reported_entry)
# Drop rows where clean_problem_reported_entry returned None
df = df.dropna(subset=['temp_problem_reported']).copy()
filtered_count = initial_rows_count - len(df)
if filtered_count > 0:
    print(f"Filtered out {filtered_count} rows due to 'problem_reported' cleaning. Remaining rows: {len(df)}")
df[PROBLEM_COLUMN] = df['temp_problem_reported']
df = df.drop(columns=['temp_problem_reported'])


# Apply comprehensive cleaning to DecodedBody to create Signature_Remover
initial_rows_count = len(df)
df['Signature_Remover'] = df[MAIL_BODY_COLUMN].apply(clean_decoded_body)
# Drop rows where Signature_Remover became empty after cleaning
df = df.dropna(subset=['Signature_Remover']).copy()
filtered_count = initial_rows_count - len(df)
if filtered_count > 0:
    print(f"Filtered out {filtered_count} rows due to 'Signature_Remover' becoming empty after cleaning. Remaining rows: {len(df)}")


# Final check to ensure Signature_Remover is not empty after all operations
initial_rows_count = len(df)
df = df[df['Signature_Remover'].astype(str).str.strip() != ''].copy()
filtered_count = initial_rows_count - len(df)
if filtered_count > 0:
    print(f"Filtered out {filtered_count} rows due to 'Signature_Remover' being empty after final strip. Remaining rows: {len(df)}")


print(f"Cleaning complete. Remaining rows: {len(df)}")

if not df.empty:
    # --- Display the first few rows of the cleaned DataFrame ---
    print(f"\n--- Head of Cleaned DataFrame ({CLEANED_OUTPUT_FILE}) ---")
    # Define key columns for preview
    preview_cols = [
        'docket_no',
        'problem_reported',
        'priority_name',
        'disposition_name',
        'sub_disposition_name',
        'DecodedBody', # Show original DecodedBody
        'Signature_Remover' # Show cleaned DecodedBody
    ]

    # Filter to only show columns that actually exist in the cleaned DataFrame
    existing_preview_cols = [col for col in preview_cols if col in df.columns]

    if existing_preview_cols:
        print(df[existing_preview_cols].head().to_markdown(index=False))
    else:
         print(df.head().to_markdown(index=False)) # Fallback if none of the key columns exist
    print("\n(Note: Full content for all columns, including DecodedBody, is in the CSV file. Enable 'Wrap Text' in spreadsheet for full view.)")


    # --- Save the cleaned DataFrame to a new CSV file ---
    print(f"\nSaving cleaned DataFrame to '{CLEANED_OUTPUT_FILE}'...")
    try:
        df.to_csv(CLEANED_OUTPUT_FILE, index=False)
        print(f"Results saved to '{CLEANED_OUTPUT_FILE}'.")
    except Exception as e:
        print(f"Error saving cleaned DataFrame to CSV: {e}", file=sys.stderr)
else:
    print("DataFrame was empty after cleaning. No output file generated.")
