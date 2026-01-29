import pdfplumber
import pandas as pd
import os
import glob

# ==========================================
# CONFIGURATION
# ==========================================
PDF_FOLDER = r"Location of downloaded PDFs" # Folder containing PDFs to process
OUTPUT_FOLDER = r"Storage for Excel Files" # Folder to save output Excel files

# Ensure the output directory exists
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

# ==========================================
# HELPER FUNCTIONS
# ==========================================

def clean_and_normalize(table):
    """
    Converts a raw list-of-lists into a pandas DataFrame.
    Fills None/NaN values with empty strings to ensure clean Excel output.
    """
    df = pd.DataFrame(table)
    df = df.fillna("")
    return df

def are_tables_mergeable(prev_df, curr_df, prev_page, curr_page):
    """
    Determines if two tables should be merged based on page continuity 
    and column structure.
    
    Parameters:
        prev_df (DataFrame): The table from the previous page
        curr_df (DataFrame): The table from the current page
        prev_page (int): Page number of the previous table
        curr_page (int): Page number of the current table
        
    Returns:
        bool: True if tables should be stitched, False otherwise.
    """
    # Criterion 1: Tables must be on consecutive pages
    if curr_page != prev_page + 1:
        return False
        
    # Criterion 2: Tables must have the same number of columns
    if len(prev_df.columns) != len(curr_df.columns):
        return False
        
    return True

# ==========================================
# MAIN PROCESSING FUNCTION
# ==========================================

def process_pdfs_with_stitching(pdf_folder, output_folder):
    # Retrieve all PDF files from the input folder
    pdf_files = glob.glob(os.path.join(pdf_folder, "*.pdf"))
    
    if not pdf_files:
        print(f"No PDF files found in directory: {pdf_folder}")
        return

    print(f"Found {len(pdf_files)} PDFs. Starting extraction process...")

    for file_path in pdf_files:
        filename = os.path.basename(file_path)
        base_name = filename.replace(".pdf", "")
        print(f"Processing: {filename}")
        
        # -------------------------------------------------------
        # STEP 1: RAW EXTRACTION
        # Extract all tables from the PDF and store them with their page numbers.
        # -------------------------------------------------------
        extracted_pieces = [] 
        
        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    # Extract tables using default settings
                    # Note: Settings can be tuned here (e.g., snap_tolerance) if needed later
                    tables = page.extract_tables()
                    
                    for table in tables:
                        df = clean_and_normalize(table)
                        extracted_pieces.append({'page': page.page_number, 'df': df})
        except Exception as e:
            print(f"Error reading PDF {filename}: {e}")
            continue
        
        if not extracted_pieces:
            print(f"No tables found in {filename}.")
            continue

        # -------------------------------------------------------
        # STEP 2: STITCHING LOGIC
        # Iterate through extracted tables and merge them if they span pages.
        # -------------------------------------------------------
        final_tables = []
        
        if extracted_pieces:
            # Initialize with the first table found
            current_merged = extracted_pieces[0]
            
            for i in range(1, len(extracted_pieces)):
                next_piece = extracted_pieces[i]
                
                # Check if the current table and the next table are part of the same dataset
                if are_tables_mergeable(current_merged['df'], next_piece['df'], 
                                        current_merged['page'], next_piece['page']):
                    
                    print(f"   Stitching table from Page {current_merged['page']} with Page {next_piece['page']}")
                    
                    df_prev = current_merged['df']
                    df_next = next_piece['df']
                    
                    # Merge the DataFrames vertically
                    # ignore_index=True resets the row numbers so they flow continuously
                    merged_df = pd.concat([df_prev, df_next], ignore_index=True)
                    
                    # Update 'current_merged' to include the newly merged data
                    # We update the page number to the 'next' page to allow for multi-page chaining
                    current_merged = {'page': next_piece['page'], 'df': merged_df}
                    
                else:
                    # If criteria not met, finalize the current table and start a new block
                    final_tables.append(current_merged['df'])
                    current_merged = next_piece
            
            # Append the final table remaining in the buffer
            final_tables.append(current_merged['df'])

        # -------------------------------------------------------
        # STEP 3: SAVE TO EXCEL
        # Write the processed tables to an Excel file with distinct sheets.
        # -------------------------------------------------------
        save_path = os.path.join(output_folder, f"{base_name}_STITCHED.xlsx")
        
        try:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                for i, df in enumerate(final_tables):
                    # Naming sheets sequentially
                    sheet_name = f"Table_{i+1}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                    
            print(f"   Successfully saved {len(final_tables)} tables to {save_path}")
            
        except Exception as e:
            print(f"   Error saving Excel file for {filename}: {e}")

    print("Batch processing complete.")

# ==========================================
# EXECUTION
# ==========================================
if __name__ == "__main__":
    process_pdfs_with_stitching(PDF_FOLDER, OUTPUT_FOLDER)