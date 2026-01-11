# /pod-check - POD Presence Validator

Compare scanned POD PDFs against manifest Excel to identify missing, present, and extra PODs. Optionally extract and compare PDF content against Excel Master Data.

## Usage

```
/pod-check [pod_folder] [manifest_path] [--output output_folder] [--compare-content] [--no-ocr] [--format FORMAT]
```

## Arguments

- `pod_folder` - Path to folder containing scanned POD PDFs (optional, uses config default)
- `manifest_path` - Path to manifest Excel file (optional, uses config default)
- `--output, -o` - Output folder for report (optional)
- `--compare-content, -c` - Extract PDF content and compare against manifest data
- `--no-ocr` - Disable OCR (use text extraction only, for text-based PDFs)
- `--format, -f` - Output format (default: md)
  - `md` - Markdown (default)
  - `csv` - CSV spreadsheet
  - `xlsx` - Excel workbook
  - `html` - HTML webpage
  - `pdf` - PDF document (requires weasyprint)
  - `all` - Export all formats at once

## Examples

```bash
# Use default paths, output as Markdown (default)
/pod-check

# Specify folder and manifest
/pod-check "D:\PODs\January2024" "D:\Data\manifest_20240115.xlsx"

# Compare PDF content with Excel Master Data
/pod-check --compare-content

# Compare content without OCR (for text-based PDFs)
/pod-check --compare-content --no-ocr

# Export as CSV
/pod-check --format csv

# Export as Excel
/pod-check --format xlsx

# Export as HTML (styled webpage)
/pod-check --format html

# Export all formats at once
/pod-check --compare-content --format all

# Full example with all options
/pod-check "D:\PODs" "D:\Data\manifest.xlsx" -c --no-ocr -f all -o "D:\Reports"
```

## Output

### Basic Mode (default)
Generates Excel report with:
- **Summary section**: Total, Present, Missing, Extra counts
- **Detail section**: Each delivery with status (Present/Missing/Extra)

Basic Columns:
- Delivery ID
- Status (Present/Missing/Extra)
- Filename (if present)
- Manifest Date
- Customer Name

### Content Comparison Mode (--compare-content)
Additional columns when content comparison is enabled:

| Column | Description |
|--------|-------------|
| Content Match | Yes / Partial / No / Error |
| Match Score | Percentage (0-100%) |
| PDF Date | Date extracted from PDF |
| Date Match | Yes / No |
| PDF Customer | Customer name found in PDF |
| Customer Match | Yes / No |
| Issues | List of discrepancies found |

Additional Summary Statistics:
- Full Match count
- Partial Match count
- No Match count
- Errors count
- Content Match Rate percentage

## Prerequisites for Content Comparison

For scanned PDFs (OCR mode):
1. Install Tesseract OCR:
   - Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki
   - Add to system PATH
2. Install Python dependencies:
   ```bash
   pip install pytesseract pdf2image Pillow
   ```

For text-based PDFs (--no-ocr):
- No additional dependencies required

## Time Saved

| Task | Manual | With Skill |
|------|--------|------------|
| Cross-check 350 PODs | 45-60 min | 30 seconds |
| Identify missing | 15-20 min | Instant |
| Compare PDF content | 2-3 hours | 5-10 min |
| Generate report | 10-15 min | Instant |

## Implementation

Run the Python script:
```bash
# Basic presence check
python skills/pod-check/pod_check.py [pod_folder] [manifest_path]

# With content comparison
python skills/pod-check/pod_check.py [pod_folder] [manifest_path] --compare-content

# Content comparison without OCR
python skills/pod-check/pod_check.py [pod_folder] [manifest_path] --compare-content --no-ocr
```
