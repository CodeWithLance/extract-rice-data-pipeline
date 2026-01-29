import pandas as pd
import os
import glob

# ==========================================
# CONFIGURATION
# ==========================================
INPUT_FOLDER = r"Files to Split" # Folder containing stitched Excel files
OUTPUT_FOLDER = r"Folder for Split Files" # Folder to save split Excel files

if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

# ==========================================
# HELPER: FIND TRUE TABLE START
# ==========================================

def find_true_table_start(df, commodity_idx):
    """
    Given the index of a 'Commodity' row, scan upwards to find the 
    actual start of the table (e.g., 'PSD Table', 'Trade Matrix').
    """
    # Don't look back further than 15 rows (efficiency & safety)
    limit = max(0, commodity_idx - 15)
    
    # We scan backwards from the commodity row
    for i in range(commodity_idx, limit - 1, -1):
        row_text = df.iloc[i].astype(str).str.cat(sep=" ").lower()
        
        # KEYWORDS: These mark the definite start of a new table
        if "psd table" in row_text:
            return i
        if "trade matrix" in row_text:
            return i
        if "prices table" in row_text:
            return i
            
    # FALLBACK: If no keyword found, assume standard header height (Country + Commodity)
    # usually 2 rows above commodity is 'Country'
    return max(0, commodity_idx - 2)

# ==========================================
# CORE PROCESSING
# ==========================================

def split_mixed_tables(df):
    split_tables = []
    
    # 1. FIND ALL COMMODITY ROWS (The Anchors)
    mask = df.astype(str).apply(
        lambda x: x.str.contains('Commodity', case=False, na=False)
    ).any(axis=1)
    
    commodity_indices = df.index[mask].tolist()
    
    # If no commodity found, return whole table
    if not commodity_indices:
        return [df]

    # 2. CALCULATE EXACT CUT POINTS
    # We want to find the "Start Row" for every detected table.
    table_starts = []
    
    for idx in commodity_indices:
        true_start = find_true_table_start(df, idx)
        table_starts.append(true_start)
        
    # 3. SLICE THE DATAFRAME
    # Logic: Table N goes from [Start_N] to [Start_N+1]
    
    for i in range(len(table_starts)):
        # Start of current table
        current_start = table_starts[i]
        
        # End of current table is the Start of the NEXT table
        if i + 1 < len(table_starts):
            # The next table starts at table_starts[i+1]
            # So this table ends right before it.
            current_end = table_starts[i+1]
        else:
            # If it's the last table, go to the very end of the dataframe
            current_end = len(df)
            
        # HANDLE THE FIRST TABLE SPECIAL CASE
        # If the first 'Commodity' appeared in the middle of the page (e.g. Row 50),
        # but there was valid data above it (Rows 0-49) that didn't belong to a previous 'Commodity'
        # (e.g. Table 1 continuation), we need to capture that.
        
        # However, usually, if i==0, we want to grab everything from 0 IF the detected start is > 0.
        # But for 'mixed' tables, usually Row 0 is the start of the first table anyway.
        if i == 0:
            current_start = 0 
            
        # SLICE
        sub_table = df.iloc[current_start:current_end].reset_index(drop=True)
        split_tables.append(sub_table)

    return split_tables

def process_splitting_workflow(input_folder, output_folder):
    files = glob.glob(os.path.join(input_folder, "*_STITCHED.xlsx"))
    
    if not files:
        print(f"No stitched Excel files found in {input_folder}.")
        return

    print(f"Starting Smart Split process on {len(files)} files...")

    for file_path in files:
        filename = os.path.basename(file_path)
        base_name = filename.replace("_STITCHED.xlsx", "")
        print(f"Processing: {filename}")
        
        try:
            xls = pd.ExcelFile(file_path)
            all_separated_tables = []
            
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                
                extracted_tables = split_mixed_tables(df)
                
                if len(extracted_tables) == 1:
                    all_separated_tables.append((sheet_name, extracted_tables[0]))
                else:
                    for idx, table in enumerate(extracted_tables):
                        new_sheet_name = f"{sheet_name}_P{idx+1}"
                        all_separated_tables.append((new_sheet_name, table))
            
            # SAVE
            save_name = f"{base_name}_SPLIT.xlsx"
            save_path = os.path.join(output_folder, save_name)
            
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                for sheet_name, df in all_separated_tables:
                    # Truncate sheet names to 31 chars for Excel compatibility
                    safe_sheet_name = sheet_name[:31] 
                    df.to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False)
            
            print(f"   Saved split tables to {save_name}")

        except Exception as e:
            print(f"   Error processing {filename}: {e}")

    print("Batch processing complete.")

if __name__ == "__main__":
    process_splitting_workflow(INPUT_FOLDER, OUTPUT_FOLDER)