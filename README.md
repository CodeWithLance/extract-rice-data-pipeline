# Philippine Rice Price Data Pipeline

This repository contains a Python automation pipeline designed to scrape, extract, and clean agricultural data from USDA Foreign Agricultural Service (FAS) PDF reports. The goal is to generate analysis-ready datasets for Philippine Rice Price movement studies.

## 📂 Project Structure

This project follows a strict hierarchy to separate Code (tracked in Git) from Data (ignored for privacy and size).

```text
DLU W1/ (Root)
├── Python Scripts/
│   ├── PDF Downloader Script/
│   │   ├── PDFDownloader.py      # Batch downloads PDFs from links.txt
│   │   └── links.txt             # List of USDA Report URLs
│   ├── PDF Table Scraper Script/
│   │   └── PDFScraper.py         # Extracts & stitches tables across pages
│   ├── Excel Splitter Script/
│   │   └── TableSplitter.py      # Separates mixed crop tables (Rice/Corn/Wheat)
│   └── Rice Filter Script/
│       └── RiceFilter.py         # Strictly filters for Rice data only
│
├── USDA FAS Files/ (Data - Not tracked in Git)
│   ├── Annual Reports/           # Raw PDF storage
│   └── Raw Data/                 # Processing outputs (Stitched -> Split -> Filtered)
│
├── requirements.txt              # Project dependencies
└── README.md                     # Project documentation
