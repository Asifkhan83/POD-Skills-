# /pod-issues - POD Quality Checker

Detect common POD issues like date mismatch, stamp problems, and customer name mismatch.

## Usage

```
/pod-issues [pod_folder] [manifest_path] [--output output_folder]
```

## Arguments

- `pod_folder` - Path to folder containing POD PDFs
- `manifest_path` - Path to manifest Excel with expected data
- `--output` - Output folder for report
- `--ocr` - Enable OCR for scanned PDFs (requires Tesseract)

## Examples

```bash
# Basic issue detection
/pod-issues "D:\PODs\January2024" "D:\Data\manifest.xlsx"

# With OCR enabled
/pod-issues "D:\PODs\January2024" "D:\Data\manifest.xlsx" --ocr

# Specify output
/pod-issues "D:\PODs\January2024" "D:\Data\manifest.xlsx" --output "D:\Reports"
```

## Issue Detection

1. **Date Issue**: Extract date from PDF, compare with manifest date (tolerance: ±2 days)
2. **Stamp Detection**: Check for stamp/signature presence in PDF
3. **Customer Mismatch**: Compare customer name in PDF with manifest (fuzzy match 80%+)
4. **Missing Signature**: Detect if signature area is empty

## Output

Generates Excel report with:
- **Summary section**: Total checked, Issues found by type
- **Detail section**: Each POD with detected issues

Columns:
- Delivery ID
- Issue Type (Date/Stamp/Customer/Signature)
- Severity (High/Medium/Low)
- Details
- Manifest Value
- PDF Value
- Needs Action (Yes/No)

## Time Saved

| Task | Manual | With Skill |
|------|--------|------------|
| Inspect 350 PDFs | 60-90 min | 2-3 min |
| Identify issues | 30-45 min | Instant |
| Categorize issues | 15-20 min | Instant |

## Dependencies

- `pdfplumber` - PDF text extraction
- `fuzzywuzzy` - Fuzzy string matching
- `pytesseract` (optional) - OCR for scanned PDFs

## Implementation

```bash
python skills/pod-issues/pod_issues.py [pod_folder] [manifest_path]
```
