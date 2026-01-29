import pandas as pd
import os
import glob

# ==========================================
# CONFIGURATION
# ==========================================
INPUT_FOLDER = r"Location of Split Excel Files" # Folder containing split Excel files
OUTPUT_FOLDER = r"Folder to Save Rice Tables Only" # Folder to save filtered Rice Excel files

if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

# ==========================================
# VALIDATION LOGIC
# ==========================================

def is_narrative_text(df):
    """
    Detects if a DataFrame is actually just a paragraph of text.
    Returns True if it looks like a sentence/story.
    """
    full_text = df.astype(str).to_string().lower()
    
    # 1. STOP WORDS
    # I removed "in", "of", "to" because valid tables use them (e.g., "Prices in Pesos")
    # We only look for "story-telling" words now.
    stop_words = [" the ", " is ", " that ", " are ", " with ", " they ", " was "]
    
    stop_word_count = 0
    for word in stop_words:
        if word in full_text:
            stop_word_count += 1
            
    if stop_word_count >= 2:
        return True

    return False

def is_valid_data_structure(df):
    """
    Checks if the physical structure looks like a table.
    """
    # 1. CELL LENGTH
    # Relaxed limit from 60 to 100 to allow for long headers
    df_str = df.astype(str)
    subset = df_str.iloc[:, :5]
    for col in subset.columns:
        if subset[col].str.len().max() > 100:
            return False 

    # 2. DIGIT DENSITY
    # Must have at least 2% numbers
    text_blob = df_str.to_string()
    total_chars = len(text_blob)
    digit_count = sum(c.isdigit() for c in text_blob)
    
    if total_chars == 0: return False
    
    density = digit_count / total_chars
    if density < 0.02:
        return False

    return True

def is_strictly_rice_table(df):
    """
    Master Validator: PRIORITIZES labels over structure.
    """
    # Grab the top 20 rows (Headers often span multiple lines)
    header_chunk = df.head(20).astype(str).to_string().lower()
    
    # === PRIORITY 1: TRUST THE LABEL ===
    # If the table explicitly says "Commodity: Rice", we KEEP it immediately.
    # We do NOT run structure checks on it, because we trust the label.
    if "commodity" in header_chunk:
        # We search line by line to ensure "Rice" is associated with "Commodity"
        for i in range(min(20, len(df))):
            row_text = df.iloc[i].astype(str).str.cat(sep=" ").lower()
            
            if "commodity" in row_text:
                if "rice" in row_text:
                    return True  # explicit MATCH -> KEEP
                elif "corn" in row_text or "wheat" in row_text:
                    return False # explicit MISMATCH -> DELETE
    
    # === PRIORITY 2: INSPECT UNLABELED DATA ===
    # If there is no "Commodity" label, it might be a text block or a messy table.
    # Now we apply the Strict Checks.
    
    # Check 1: Is it a story?
    if is_narrative_text(df):
        return False

    # Check 2: Is it structurally garbage?
    if not is_valid_data_structure(df):
        return False

    # Check 3: Content Fallback
    # It passed structure checks, now does it at least mention Rice?
    has_rice = "rice" in header_chunk
    has_corn = "corn" in header_chunk
    has_wheat = "wheat" in header_chunk
    
    if has_rice and not (has_corn or has_wheat):
        return True
        
    return False

# ==========================================
# MAIN EXECUTION
# ==========================================

def process_rice_filtering(input_folder, output_folder):
    files = glob.glob(os.path.join(input_folder, "*_SPLIT.xlsx"))
    
    if not files:
        print(f"No files found in {input_folder}.")
        return

    print(f"Starting Smart Filter (Label-First Mode) on {len(files)} files...")

    for file_path in files:
        filename = os.path.basename(file_path)
        base_name = filename.replace("_SPLIT.xlsx", "")
        print(f"Scanning: {filename}")
        
        try:
            xls = pd.ExcelFile(file_path)
            rice_sheets = [] 
            
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                
                if is_strictly_rice_table(df):
                    rice_sheets.append((sheet_name, df))
            
            if rice_sheets:
                save_name = f"{base_name}_RICE.xlsx"
                save_path = os.path.join(output_folder, save_name)
                
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    for sheet_name, df in rice_sheets:
                        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                
                print(f"   [MATCH] Saved {len(rice_sheets)} verified Rice tables.")
            else:
                print(f"   [SKIP] No Rice tables found in {filename}")

        except Exception as e:
            print(f"   [ERROR] Could not process {filename}: {e}")

    print("Filtering complete.")

if __name__ == "__main__":
    process_rice_filtering(INPUT_FOLDER, OUTPUT_FOLDER)