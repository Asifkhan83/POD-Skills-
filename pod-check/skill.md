# /pod-check - POD Presence Validator

Compare scanned POD PDFs against manifest Excel to identify missing, present, and extra PODs.

## Usage

```
/pod-check [pod_folder] [manifest_path] [--output output_folder]
```

## Arguments

- `pod_folder` - Path to folder containing scanned POD PDFs (optional, uses config default)
- `manifest_path` - Path to manifest Excel file (optional, uses config default)
- `--output` - Output folder for report (optional)

## Examples

```bash
# Use default paths from config
/pod-check

# Specify folder and manifest
/pod-check "D:\PODs\January2024" "D:\Data\manifest_20240115.xlsx"

# Specify all paths
/pod-check "D:\PODs\January2024" "D:\Data\manifest.xlsx" --output "D:\Reports"
```

## Output

Generates Excel report with:
- **Summary section**: Total, Present, Missing, Extra counts
- **Detail section**: Each delivery with status (Present/Missing/Extra)

Columns:
- Delivery ID
- Status (Present/Missing/Extra)
- Filename (if present)
- Manifest Date
- Customer Name

## Time Saved

| Task | Manual | With Skill |
|------|--------|------------|
| Cross-check 350 PODs | 45-60 min | 30 seconds |
| Identify missing | 15-20 min | Instant |
| Generate report | 10-15 min | Instant |

## Implementation

Run the Python script:
```bash
python skills/pod-check/pod_check.py [pod_folder] [manifest_path]
```
