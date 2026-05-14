# Casamia-Internship-Projects
This repository contains all the tasks performed at Casamia during my internship which include all the programming, automation and extraction tasks.

# Showroom & Catalog PDF Extraction Suite

This repository contains automated Python scripts designed to extract structured inventory data (SKUs, Item Codes, Descriptions, Brands) from multi-format PDF catalogs and showroom inventories.

The extraction pipeline handles highly irregular layouts using native table parsing via PyMuPDF, sliding-window regex processing, and Tesseract OCR fallback for pages lacking readable text layers. Output is compiled into auto-fitted, styled Excel spreadsheets.

## File Matrix & Specifications

| Script Name | Target PDF | Output Excel | Core Architecture & Features |
| :--- | :--- | :--- | :--- |
| **`accessories1.py`** | `ACCESSORIES SHOWROOM-1-45.pdf` | `Extracted_Showroom_Complete_Data.xlsx` | **Pure PyMuPDF Tables:** Tracks persistent `DOC NO` markers. Applies an 80+ word exclusion array to filter noise and handles Column 0/1 floating-point SKU shifts. |
| **`accessories46.py`** | `ACCESSORIES SHOWROOM-46-82.pdf` | `Extracted_Showroom_Accessories_Data.xlsx` | **Hybrid Engine:** Prioritizes native matrix parsing; falls back to sliding-window text parsing and Tesseract OCR. Maps default brand context to "Casamia Showroom". |
| **`extract_catalog.py`** | `OUTDOOR ACCESSORIES.pdf` | `Extracted_OUTDOOR ACCESSORIES_Catalog_Data.xlsx` | **Catalog Optimized:** Targets key-value text pointers for outdoor planters. Includes OCR fallback and internal string scanning to detect embedded designer brands (e.g., Vondom, Serralunga). |
| **`extract_table.py`** | `GF FERRE HOME-ACCESSORIES.pdf` | `Extracted_GF_FERRE_Accessories_Data.xlsx` | **Brand Optimized:** Parses strict multi-dot corporate SKUs (`X.XXX.XXX.XXX`), maps tabular matrices, and assigns specific brands dynamically based on body text context. |


# Product Management & PDF Organization Scripts

This repository contains two automation scripts designed for inventory master data validation and bulk PDF asset organization.

## Repository Contents

### 1. `tsk.py` — Item Master Data Validation
Validates actual files present in a folder against product entries listed in an Excel Master sheet.
* **Brand Filtering:** Isolates expected records dynamically based on user input.
* **Gap Analysis:** Identifies products missing from the folder and unmapped files.
* **Metadata Extraction:** Scans file content via regex to verify internal reference strings against master spreadsheet values.
* **Output:** Generates a complete status audit trail saved as `[Brand]_validation_report.csv`.

### 2. `pdf.py` — Automated PDF Asset Routing
Scans nested directories for PDF assets and organizes them into structured brand folders based on an Excel mapping dictionary.
* **Intelligent Mapping:** Links file names directly to product codes and brand identifiers in the master sheet.
* **Automated File Routing:** Generates brand-specific target directories on the fly and securely duplicates matching PDFs.
* **Audit Tracking:** Outputs an execution report documenting total successful file transfers per brand (`brand_match_summary.csv`).


# Automation & Web Application Scripts

This repository contains two backend scripts: an automated Excel processing utility designed for dynamic image embedding, and a lightweight Flask web application framework.

## File Overview

| Script Name | Environment | Purpose / Features |
| :--- | :--- | :--- |
| **`main.py`** | Desktop / File System | **Batch Excel Automation:** Iterates through Excel documents to dynamically embed product images based on SKUs/Item IDs. Handles cell resizing and automated placeholder injection for missing assets. |
| **`app.py`** | Web Server | **Web Framework:** A boilerplate Flask application configured for rapid web prototyping and deployment. |

## Script Specifications

### 1. `main.py` — Excel Asset Injection
Automates the integration of external product imagery directly into standard Excel inventory or quotation spreadsheets using `openpyxl`.

* **Dynamic ID Mapping:** Scans Column A for product identification codes and queries the designated `IMAGE_FOLDER` for corresponding `.jpeg` assets formatted as `#[item_id].jpeg`.
* **Layout Optimization:** Standardizes document layouts by setting explicit row heights (`100`) and column widths (`50`) to ensure images align perfectly within grid boundaries.
* **Error Handling & Placeholders:** Implements automated fallback routing. If a specific product image is absent or throws an OS rendering error, the script intercepts the exception and cleanly places a default placeholder graphic (`PLACEHOLDER_IMAGE_LOCATION`).
* **Batch Processing:** Sequentially targets all `.xlsx` files inside the input directory and generates preserved duplicate artifacts prefixed with `Out `.

### 2. `app.py` — Flask Microservice
Establishes a foundational backend gateway using the Flask micro-framework.

* **Routing:** Configures the root route (`/`) to output a verification string confirming active server operational status.
* **Development Mode:** Operates with integrated `debug=True` settings enabled to facilitate hot-reloading and trace logging during development.
