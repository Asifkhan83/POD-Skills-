# /pod-archive - POD Archive Manager

Organize and archive processed PODs by date, customer, or status.

## Usage

```
/pod-archive [source_folder] [archive_folder] [--mode by-date|by-customer|by-status] [--manifest path]
```

## Arguments

- `source_folder` - Path to folder with PODs to archive
- `archive_folder` - Destination archive folder
- `--mode` - Organization mode: by-date, by-customer, by-status (default: by-date)
- `--manifest` - Path to manifest Excel for metadata
- `--status-report` - Path to status report for status-based archiving
- `--copy` - Copy files instead of moving (default: move)
- `--dry-run` - Preview changes without moving files

## Examples

```bash
# Archive by date (default)
/pod-archive "D:\PODs\Processed" "D:\Archive"

# Archive by customer
/pod-archive "D:\PODs\Processed" "D:\Archive" --mode by-customer --manifest "D:\Data\manifest.xlsx"

# Archive by status
/pod-archive "D:\PODs\Processed" "D:\Archive" --mode by-status --status-report "D:\Reports\status.xlsx"

# Dry run to preview
/pod-archive "D:\PODs\Processed" "D:\Archive" --dry-run
```

## Archive Modes

1. **by-date**: `Archive/2024/01/15/` - organize by delivery date
2. **by-customer**: `Archive/CustomerABC/2024-01/` - group by customer
3. **by-status**: `Archive/Completed/`, `Archive/Issues/` - separate by resolution

## Output

- Organized folder structure in archive destination
- Archive log Excel with:
  - Delivery ID
  - Source Path
  - Destination Path
  - Archive Date
  - Mode Used

## Time Saved

| Task | Manual | With Skill |
|------|--------|------------|
| Organize 350 files | 25-35 min | < 1 min |
| Create folder structure | 10-15 min | Instant |
| Move files | 15-20 min | Instant |

## Implementation

```bash
python skills/pod-archive/pod_archive.py [source] [archive] [--mode mode]
```
