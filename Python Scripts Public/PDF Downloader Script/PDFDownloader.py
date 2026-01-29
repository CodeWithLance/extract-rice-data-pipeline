import requests
import time
import urllib.parse
import re
import os

# --- CONFIGURATION ---
# 1. Create a "pdfs" folder automatically inside your project folder
DOWNLOAD_FOLDER = "Location of downloaded PDFs" # Folder to save downloaded PDFs

# 2. Paste your USDA links here
# Read links from text file instead of hardcoding
with open("links.txt", "r") as f:
    links = [line.strip() for line in f if line.strip()]

def setup_folder():
    """Creates the download folder if it doesn't exist."""
    if not os.path.exists(DOWNLOAD_FOLDER):
        os.makedirs(DOWNLOAD_FOLDER)
        print(f"Created folder: {os.path.abspath(DOWNLOAD_FOLDER)}")
    else:
        print(f"Saving to: {os.path.abspath(DOWNLOAD_FOLDER)}")

def get_filename_from_cd(cd):
    """Get filename from content-disposition header."""
    if not cd:
        return None
    fname = re.findall('filename=(.+)', cd)
    if len(fname) == 0:
        return None
    return fname[0].strip('"')

def download_usda_pdf(url):
    try:
        # Send request
        response = requests.get(url, stream=True)
        response.raise_for_status()
        
        # Determine Filename (Header -> URL -> Default)
        filename = get_filename_from_cd(response.headers.get('content-disposition'))
        
        if not filename and "filename=" in url:
            parsed = urllib.parse.urlparse(url)
            query_params = urllib.parse.parse_qs(parsed.query)
            filename = query_params.get('filename', [None])[0]
            
        if not filename:
            filename = f"usda_report_{int(time.time())}.pdf" # Fallback timestamp name

        # Clean filename
        filename = urllib.parse.unquote(filename)
        
        # Save the file
        file_path = os.path.join(DOWNLOAD_FOLDER, filename)
        
        # Check if file already exists to avoid re-downloading
        if os.path.exists(file_path):
            print(f"Skipping (already exists): {filename}")
            return

        with open(file_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
                
        print(f"Downloaded: {filename}")
        
    except Exception as e:
        print(f"Failed to download {url}: {e}")

# --- MAIN EXECUTION ---
if __name__ == "__main__":
    setup_folder()
    print(f"Starting download of {len(links)} files...\n")

    for index, link in enumerate(links):
        download_usda_pdf(link)
        
        # Simple progress indicator
        print(f"   [Processed {index + 1} of {len(links)}]")
        
        # Polite pause
        time.sleep(2) 

    print("\nJob Complete!")