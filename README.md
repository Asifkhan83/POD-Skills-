# POD Management Skills

A suite of 5 Python-based skills to automate Proof of Delivery (POD) management workflow. Designed for handling 200-500 PODs/day with Excel reports as output.

## Demo Video

[![POD Skills Demo](https://img.youtube.com/vi/8K_Lio18udA/maxresdefault.jpg)](https://www.youtube.com/watch?v=8K_Lio18udA)

> Click the image above to watch the demo video

---

## Quick Start

### Installation

```bash
# Install dependencies
pip install pandas openpyxl pdfplumber fuzzywuzzy python-Levenshtein
```

### Configuration

Edit `shared/config.py` to set your default paths:
```python
DEFAULT_POD_FOLDER = r"D:\PODs"
DEFAULT_MANIFEST_PATH = r"D:\Data\manifest.xlsx"
DEFAULT_OUTPUT_FOLDER = r"D:\Reports"
```

---

## Skills Overview

| Skill | Purpose | Daily Time Saved | Mental Effort Reduced |
|-------|---------|------------------|----------------------|
| `/pod-check` | Compare PDFs vs manifest | 45-60 min | Cross-referencing 500 items |
| `/pod-status` | Track closure readiness | 20-30 min | Status consolidation |
| `/pod-issues` | Detect quality issues | 60-90 min | Visual PDF inspection |
| `/pod-archive` | Organize & archive PODs | 25-35 min | File organization |
| `/pod-email` | Generate issue emails | 20-30 min | Email composition |
| **TOTAL** | | **170-245 min/day** | **Significant cognitive load** |

**Weekly savings: 14-20 hours of manual work**

---

## Skill 1: `/pod-check` - POD Presence Validator

### What it replaces
- Manual checking of 200-500 filenames against Excel manifest
- Visual scanning of folder contents
- Cross-referencing delivery numbers one by one

### Usage
```bash
python skills/pod-check/pod_check.py "D:\PODs\January" "D:\Data\manifest.xlsx"
```

### Output
Excel report with:
- Summary: Total, Present, Missing, Extra counts
- Details: Each delivery with status (Present/Missing/Extra)

### Time Comparison
| Task | Manual | With Skill |
|------|--------|------------|
| Cross-check 350 PODs | 45-60 min | 30 seconds |
| Identify missing | 15-20 min | Instant |
| Generate report | 10-15 min | Instant |

---

## Skill 2: `/pod-status` - POD Status Tracker

### What it replaces
- Manual status tracking in Excel
- Determining which deliveries are complete
- Identifying what's blocking closure

### Usage
```bash
python skills/pod-status/pod_status.py "D:\Data\master.xlsx" --check-report "D:\Reports\pod_check.xlsx"
```

### Output
Excel report with:
- Summary: Total, Complete, Pending, Has Issues
- Details: Each delivery with consolidated status
- Closure-ready batch list

### Quality Improvement
- Instant visibility into closure readiness
- No missed items ready for closure
- Clear tracking of what's blocking each delivery

---

## Skill 3: `/pod-issues` - POD Quality Checker

### What it replaces
- Opening each PDF to visually inspect
- Comparing dates manually
- Checking stamp clarity
- Verifying customer names

### Usage
```bash
python skills/pod-issues/pod_issues.py "D:\PODs\January" "D:\Data\manifest.xlsx"
```

### Issue Detection
1. **Date Issue**: Extract date from PDF, compare with manifest
2. **Stamp Detection**: Check for stamp/signature presence
3. **Customer Mismatch**: Fuzzy match customer name (80%+ threshold)

### Output
Excel report with:
- Severity: High/Medium/Low
- Issue Type: Date/Stamp/Customer
- Expected vs Actual values
- Needs Action flag

### Time Comparison
| Task | Manual | With Skill |
|------|--------|------------|
| Inspect 350 PDFs | 60-90 min | 2-3 min |
| Identify issues | 30-45 min | Instant |
| Categorize issues | 15-20 min | Instant |

---

## Skill 4: `/pod-archive` - POD Archive Manager

### What it replaces
- Manual folder organization of processed PODs
- Moving files to archive locations
- Creating organized folder structures
- Cleaning up working directories

### Usage
```bash
# Archive by date
python skills/pod-archive/pod_archive.py "D:\PODs\Processed" "D:\Archive"

# Archive by customer
python skills/pod-archive/pod_archive.py "D:\PODs\Processed" "D:\Archive" --mode by-customer --manifest "D:\Data\manifest.xlsx"

# Preview without moving
python skills/pod-archive/pod_archive.py "D:\PODs\Processed" "D:\Archive" --dry-run
```

### Archive Modes
- **by-date**: `Archive/2024/01/15/`
- **by-customer**: `Archive/CustomerABC/2024-01/`
- **by-status**: `Archive/Completed/`, `Archive/Issues/`

### Output
- Organized archive folder structure
- Archive log Excel with audit trail

### Time Comparison
| Task | Manual | With Skill |
|------|--------|------------|
| Organize 350 files | 25-35 min | < 1 min |
| Create folder structure | 10-15 min | Instant |
| Move files | 15-20 min | Instant |

---

## Skill 5: `/pod-email` - POD Issue Email Generator

### What it replaces
- Manually composing emails for each issue type
- Looking up business contacts
- Formatting issue details
- Aggregating issues per business

### Usage
```bash
# Generate quality issue emails
python skills/pod-email/pod_email.py "D:\Reports\pod_issues.xlsx" --contacts "D:\Data\contacts.xlsx"

# Weekly summary
python skills/pod-email/pod_email.py "D:\Reports\pod_issues.xlsx" --template summary
```

### Email Templates
1. **missing** - Missing POD notification
2. **quality** - POD quality issues
3. **resolution** - Resolution request
4. **summary** - Weekly summary

### Output
- Email drafts as .txt files (one per business)
- Email log Excel

### Time Comparison
| Task | Manual | With Skill |
|------|--------|------------|
| Compose 10 emails | 20-30 min | < 1 min |
| Format issue details | 10-15 min | Instant |
| Aggregate by business | 5-10 min | Instant |

---

## Typical Daily Workflow

```bash
# 1. Morning: Check POD presence
python skills/pod-check/pod_check.py "D:\PODs\Today" "D:\Data\manifest.xlsx"

# 2. Identify quality issues
python skills/pod-issues/pod_issues.py "D:\PODs\Today" "D:\Data\manifest.xlsx"

# 3. Consolidate status
python skills/pod-status/pod_status.py "D:\Data\master.xlsx" \
    --check-report "D:\Reports\pod_check_report.xlsx" \
    --issues-report "D:\Reports\pod_issues_report.xlsx"

# 4. Generate emails for issues
python skills/pod-email/pod_email.py "D:\Reports\pod_issues_report.xlsx"

# 5. End of day: Archive completed PODs
python skills/pod-archive/pod_archive.py "D:\PODs\Completed" "D:\Archive" --mode by-date
```

---

## Manifest Excel Format

Required columns (names configurable in `config.py`):

| Column | Description |
|--------|-------------|
| Delivery ID | Unique delivery number (e.g., 9354302576) |
| Delivery Date | Expected delivery date |
| Customer Name | Customer/business name |
| Status | Current status (optional) |

---

## Dependencies

- `pandas` - Excel reading/writing
- `openpyxl` - Excel formatting
- `pdfplumber` - PDF text extraction
- `fuzzywuzzy` - Fuzzy string matching
- `python-Levenshtein` - Speed up fuzzy matching (optional)

---

## File Structure

```
skills/
├── pod-check/
│   ├── skill.md          # Skill documentation
│   └── pod_check.py      # Main script (~200 LOC)
├── pod-issues/
│   ├── skill.md
│   └── pod_issues.py     # PDF analysis (~400 LOC)
├── pod-status/
│   ├── skill.md
│   └── pod_status.py     # Status consolidation (~250 LOC)
├── pod-archive/
│   ├── skill.md
│   └── pod_archive.py    # Archive management (~250 LOC)
├── pod-email/
│   ├── skill.md
│   └── pod_email.py      # Email generation (~350 LOC)
├── shared/
│   ├── __init__.py
│   ├── config.py         # Configuration
│   └── excel_utils.py    # Excel utilities
└── README.md             # This file
```

---

## Customization

### Adding New Email Templates

Edit `pod_email.py` and add to the `TEMPLATES` dict:

```python
TEMPLATES['custom'] = {
    'subject': 'Custom Subject - {count} items',
    'body': """Dear {contact_name},

{issue_list}

Best regards,
Team
"""
}
```

### Changing Column Mappings

Edit `shared/config.py`:

```python
MANIFEST_COLUMNS = {
    'delivery_id': 'Your Delivery ID Column',
    'date': 'Your Date Column',
    'customer': 'Your Customer Column',
    'status': 'Your Status Column',
}
```

---

## Measurable Outcomes

### Before Automation
- 3-4 hours daily on POD management
- Manual cross-referencing prone to errors
- Delayed issue detection
- Inconsistent email communication
- Disorganized file storage

### After Automation
- 15-30 minutes daily on POD management
- 100% accuracy in presence checking
- Immediate issue detection
- Consistent, professional emails
- Organized, searchable archive

### ROI Calculation
- Time saved: 14-20 hours/week
- At $30/hour: **$420-600/week savings**
- Error reduction: Prevents missed deliveries and delayed issue resolution
- Consistency: Standardized process across all team members
