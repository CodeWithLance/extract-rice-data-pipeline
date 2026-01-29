import os
import sys
import pandas as pd
from img2table.document import PDF
from img2table.ocr import TesseractOCR

# ==========================================
# 0. WINDOWS FIX
# ==========================================
tesseract_dir = r"C:\Program Files\Tesseract-OCR"
if os.path.exists(tesseract_dir):
    os.environ["PATH"] += os.pathsep + tesseract_dir

# ==========================================
# 1. CONFIGURATION
# ==========================================
PDF_FOLDER = r"PDF Folder" # Update this path as needed
OUTPUT_FOLDER = r"Location to save output files" # Update this path as needed

TARGET_FILES = [
    "PDF Files here.pdf" # Update with actual PDF filenames
]

if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

# ==========================================
# 2. INITIALIZE ENGINE
# ==========================================
print("Initializing Tesseract OCR...")
try:
    ocr = TesseractOCR(n_threads=8, lang="eng")
except Exception as e:
    print(f"[FATAL] Tesseract error: {e}")
    sys.exit()

# ==========================================
# 3. HELPER: GARBAGE FILTER
# ==========================================
def is_valid_table(df):
    if df.shape[0] < 2 or df.shape[1] < 2: return False
    
    text_blob = df.astype(str).to_string().replace(" ", "").replace("\n", "")
    total_chars = len(text_blob)
    if total_chars == 0: return False
    
    digit_count = sum(c.isdigit() for c in text_blob)
    return (digit_count / total_chars) >= 0.08

# ==========================================
# 4. HELPER: TABLE STITCHER (NEW!)
# ==========================================
def stitch_tables(table_list):
    """
    Merges tables that are split across consecutive pages.
    Input: List of dictionaries [{'page': 1, 'df': dataframe}, ...]
    Output: Consolidated list of dataframes
    """
    if not table_list: return []
    
    stitched = []
    
    # Start with the first table
    current_entry = table_list[0]
    
    for next_entry in table_list[1:]:
        curr_df = current_entry['df']
        next_df = next_entry['df']
        
        # CHECK 1: Are they consecutive pages?
        is_consecutive = (next_entry['page'] == current_entry['page'] + 1)
        
        # CHECK 2: Do they have the same width (column count)?
        is_same_width = (len(curr_df.columns) == len(next_df.columns))
        
        if is_consecutive and is_same_width:
            print(f"   ...Stitching Page {current_entry['page']} and {next_entry['page']}")
            
            # CHECK 3: Header Removal
            # Often the second table repeats the header. If the first row of the new table 
            # looks like the header of the old table, drop it.
            # (We convert to string to compare loosely)
            row_0_new = next_df.iloc[0].astype(str).values.tolist()
            row_0_old = curr_df.iloc[0].astype(str).values.tolist()
            
            # If >50% of the cells match, it's likely a repeated header
            match_count = sum([1 for i, x in enumerate(row_0_new) if x == row_0_old[i]])
            if match_count > (len(row_0_new) / 2):
                next_df = next_df.iloc[1:] # Drop the duplicate header
            
            # MERGE
            merged_df = pd.concat([curr_df, next_df], ignore_index=True)
            
            # Update the 'current' entry to be this new giant table
            current_entry['df'] = merged_df
            current_entry['page'] = next_entry['page'] # Update page marker to the latest one
            
        else:
            # Not compatible? Save the old one and start a new chain
            stitched.append(current_entry)
            current_entry = next_entry
            
    # Append the final straggler
    stitched.append(current_entry)
    
    return stitched

# ==========================================
# 5. MAIN EXTRACTION
# ==========================================
def process_pdf(pdf_path, filename):
    print(f"\nProcessing: {filename}...")
    doc = PDF(src=pdf_path)

    try:
        extracted_tables = doc.extract_tables(
            ocr=ocr,
            implicit_rows=True,
            borderless_tables=True,
            min_confidence=40
        )
    except Exception as e:
        print(f"   [ERROR] OCR Failed: {e}")
        return

    # A. COLLECT RAW TABLES
    raw_tables = []
    
    for page_idx, tables in extracted_tables.items():
        for table in tables:
            df = table.df
            
            # Cleanup
            df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
            df = df.apply(lambda x: x.map(lambda y: " ".join(y.split()) if isinstance(y, str) else y))

            if is_valid_table(df):
                # Store metadata for stitching
                raw_tables.append({
                    'page': page_idx + 1, # img2table pages are 0-indexed
                    'df': df
                })

    # B. STITCH THEM
    print(f"   Found {len(raw_tables)} fragments. Stitching...")
    final_tables = stitch_tables(raw_tables)
    print(f"   Final count: {len(final_tables)} tables.")

    # C. SAVE
    if final_tables:
        save_name = filename.replace(".pdf", "_OCR.xlsx")
        save_path = os.path.join(OUTPUT_FOLDER, save_name)
        
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            for idx, entry in enumerate(final_tables):
                # Naming convention: T1_PageEnd5
                sheet_name = f"Table_{idx+1}_EndsP{entry['page']}"
                entry['df'].to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        
        print(f"   ✅ Saved to {save_name}")
    else:
        print(f"   [WARNING] No data found.")

# ==========================================
# 6. EXECUTION
# ==========================================
if __name__ == "__main__":
    for pdf_file in TARGET_FILES:
        full_path = os.path.join(PDF_FOLDER, pdf_file)
        if os.path.exists(full_path):
            process_pdf(full_path, pdf_file)
        else:
            print(f"[MISSING] {full_path}")