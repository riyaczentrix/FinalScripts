import pandas as pd
import re
import nltk
from convosense_utilities.email_signature_remover import remove_sign

# Download NLTK tokenizer
nltk.download('punkt')

# Load Excel file
file_path = "ticket_details_fiftyThousand.xlsx"
df = pd.read_excel(file_path)

# Define common signature-indicating keywords (case-insensitive)
signature_keywords = [
    r'Best Regards',
    r'Warm Regards',
    r'Thanks & Regards',
    r'Thanks and Regards',
    r'Thanks',
    r'Regards',
    r'Thank you',
    r'Disclaimer'
]

# Precompile patterns
compiled_keywords = [re.compile(k, re.IGNORECASE) for k in signature_keywords]

# Preprocessing before signature removal
def preprocess_text(text):
    if pd.isnull(text):
        return ""

    # Escape backslashes
    text = text.replace("\\", "\\\\")

    # Normalize "thank you"
    text = re.sub(r'\b[Tt]hank you\b', 'Thanks', text)

    for pattern in compiled_keywords:
        match = pattern.search(text)
        if match:
            start = match.start()
            if start > 0 and text[start - 1] not in ['.', '\n']:
                text = text[:start] + '.' + text[start:]

            keyword = match.group()
            text = re.sub(
                re.escape(keyword) + r'(?!\n)',
                keyword + '\n',
                text
            )
    return text

# Apply preprocessing + signature removal
def process_email(body):
    preprocessed = preprocess_text(body)
    cleaned = remove_sign(preprocessed)
    return cleaned

# Apply to the DataFrame
df['signature_removed'] = df['DecodedBody'].apply(process_email)

# Save to new file
output_file = "ticket_details_fiftyThousand_cleaned.xlsx"
df.to_excel(output_file, index=False)

print(f"✅ Processed and saved to: {output_file}")
