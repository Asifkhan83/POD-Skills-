# /pod-status - POD Status Tracker

Consolidate POD status from multiple sources and track closure readiness.

## Usage

```
/pod-status [master_excel] [--check-report pod_check_report] [--issues-report issues_report]
```

## Arguments

- `master_excel` - Path to master tracking Excel (optional, uses config default)
- `--check-report` - Path to pod-check report for presence data
- `--issues-report` - Path to pod-issues report for issue data
- `--output` - Output folder for report

## Examples

```bash
# Basic status from master Excel
/pod-status "D:\Data\master.xlsx"

# With pod-check results
/pod-status "D:\Data\master.xlsx" --check-report "D:\Reports\pod_check_report.xlsx"

# Full consolidation
/pod-status "D:\Data\master.xlsx" --check-report "D:\Reports\pod_check.xlsx" --issues-report "D:\Reports\pod_issues.xlsx"
```

## Output

Generates Excel report with:
- **Summary section**: Total, Complete, Pending, Has Issues
- **Closure-ready batch**: List of deliveries ready to close
- **Detail section**: Each delivery with consolidated status

Columns:
- Delivery ID
- POD Received (Yes/No)
- Has Issues (Yes/No)
- Issue Details
- Resolution Status
- Ready to Close (Yes/No)

## Time Saved

| Task | Manual | With Skill |
|------|--------|------------|
| Consolidate status | 20-30 min | Instant |
| Identify closure-ready | 10-15 min | Instant |
| Generate report | 5-10 min | Instant |

## Implementation

```bash
python skills/pod-status/pod_status.py [master_excel] [options]
```
