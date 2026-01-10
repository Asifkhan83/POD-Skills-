# /pod-email - POD Issue Email Generator

Generate draft emails for POD issue communication with businesses.

## Usage

```
/pod-email [issues_report] [--contacts contacts_excel] [--template template_name]
```

## Arguments

- `issues_report` - Path to pod-issues report Excel
- `--contacts` - Path to business contacts Excel
- `--template` - Template type: missing, quality, resolution, summary (default: quality)
- `--output` - Output folder for email drafts
- `--group-by` - Group issues: by-business, by-type (default: by-business)

## Examples

```bash
# Generate emails from issues report
/pod-email "D:\Reports\pod_issues_report.xlsx"

# With contacts and template
/pod-email "D:\Reports\pod_issues.xlsx" --contacts "D:\Data\contacts.xlsx" --template missing

# Weekly summary
/pod-email "D:\Reports\pod_issues.xlsx" --template summary
```

## Templates

1. **missing** - Missing POD notification
2. **quality** - POD quality issues (date/stamp/customer)
3. **resolution** - Resolution request
4. **summary** - Weekly summary

## Contacts Excel Format

Required columns:
- Business Name
- Contact Email
- Contact Name

## Output

- Email drafts as .txt files (one per business/issue group)
- Email log Excel with:
  - Business Name
  - Subject
  - Recipients
  - Issue Count
  - Draft File Path

## Time Saved

| Task | Manual | With Skill |
|------|--------|------------|
| Compose 10 emails | 20-30 min | < 1 min |
| Format issue details | 10-15 min | Instant |
| Aggregate by business | 5-10 min | Instant |

## Implementation

```bash
python skills/pod-email/pod_email.py [issues_report] [--contacts path]
```
