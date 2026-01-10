"""
POD Issue Email Generator - /pod-email
Generate draft emails for POD issue communication.

Usage:
    python pod_email.py [issues_report] [--contacts contacts_excel] [--template name]
"""
import sys
import argparse
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional
from collections import defaultdict

import pandas as pd

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))
from shared.config import PODConfig
from shared.excel_utils import write_report


# Email templates
TEMPLATES = {
    'missing': {
        'subject': 'Action Required: Missing POD Documentation - {count} Delivery(ies)',
        'body': """Dear {contact_name},

We are reaching out regarding missing Proof of Delivery (POD) documentation for the following deliveries:

{issue_list}

Could you please provide the scanned POD documents for the above deliveries at your earliest convenience?

If you have any questions or need additional information, please don't hesitate to reach out.

Thank you for your prompt attention to this matter.

Best regards,
POD Management Team
"""
    },

    'quality': {
        'subject': 'POD Quality Issues Requiring Attention - {count} Item(s)',
        'body': """Dear {contact_name},

We have identified quality issues with the following POD documents that require your attention:

{issue_list}

Please review and provide corrected POD documentation or clarification for the above items.

Issue Details:
{issue_details}

Thank you for your cooperation.

Best regards,
POD Management Team
"""
    },

    'resolution': {
        'subject': 'POD Issue Resolution Request - {count} Outstanding Item(s)',
        'body': """Dear {contact_name},

This is a follow-up regarding outstanding POD issues that require resolution:

{issue_list}

These items have been pending resolution. Please provide an update on the status or the corrected documentation.

If you need any additional information or support, please let us know.

Thank you for your attention to this matter.

Best regards,
POD Management Team
"""
    },

    'summary': {
        'subject': 'Weekly POD Status Summary - {date}',
        'body': """Dear {contact_name},

Please find below the weekly POD status summary for your business:

Summary:
- Total Deliveries: {total}
- PODs Received: {received}
- PODs Missing: {missing}
- Issues Found: {issues}

Outstanding Items:
{issue_list}

Please address any outstanding items at your earliest convenience.

Best regards,
POD Management Team
"""
    }
}


def load_issues_report(report_path: Path) -> pd.DataFrame:
    """
    Load issues report Excel.

    Args:
        report_path: Path to issues report

    Returns:
        DataFrame with issues
    """
    # Find the header row dynamically
    df_raw = pd.read_excel(report_path, header=None)
    header_row = 0
    for i, row in df_raw.iterrows():
        if 'Delivery ID' in str(row.values):
            header_row = i
            break

    df = pd.read_excel(report_path, skiprows=header_row)
    return df


def load_contacts(contacts_path: Path) -> Dict[str, Dict]:
    """
    Load business contacts.

    Args:
        contacts_path: Path to contacts Excel

    Returns:
        Dict mapping business name to contact info
    """
    if not contacts_path or not contacts_path.exists():
        return {}

    df = pd.read_excel(contacts_path)

    contacts = {}
    for _, row in df.iterrows():
        business = str(row.get('Business Name', '')).strip()
        if business:
            contacts[business] = {
                'email': row.get('Contact Email', ''),
                'name': row.get('Contact Name', 'Team'),
            }

    return contacts


def group_issues_by_business(
    issues_df: pd.DataFrame,
    manifest_df: pd.DataFrame = None
) -> Dict[str, List[Dict]]:
    """
    Group issues by business/customer.

    Args:
        issues_df: Issues DataFrame
        manifest_df: Optional manifest for customer lookup

    Returns:
        Dict mapping business name to list of issues
    """
    grouped = defaultdict(list)

    for _, row in issues_df.iterrows():
        # Try to get business from manifest or use default
        business = "Unknown Business"

        # For now, group all under one business if no manifest
        # In production, you'd look up the customer from manifest

        grouped[business].append({
            'delivery_id': row.get('Delivery ID', ''),
            'issue_type': row.get('Issue Type', ''),
            'severity': row.get('Severity', ''),
            'details': row.get('Details', ''),
            'expected': row.get('Expected Value', ''),
            'actual': row.get('PDF Value', ''),
        })

    return dict(grouped)


def group_issues_by_type(issues_df: pd.DataFrame) -> Dict[str, List[Dict]]:
    """
    Group issues by type.

    Args:
        issues_df: Issues DataFrame

    Returns:
        Dict mapping issue type to list of issues
    """
    grouped = defaultdict(list)

    for _, row in issues_df.iterrows():
        issue_type = row.get('Issue Type', 'Unknown')

        grouped[issue_type].append({
            'delivery_id': row.get('Delivery ID', ''),
            'severity': row.get('Severity', ''),
            'details': row.get('Details', ''),
            'expected': row.get('Expected Value', ''),
            'actual': row.get('PDF Value', ''),
        })

    return dict(grouped)


def format_issue_list(issues: List[Dict], include_details: bool = False) -> str:
    """
    Format issues as bullet list.

    Args:
        issues: List of issue dicts
        include_details: Whether to include full details

    Returns:
        Formatted string
    """
    lines = []

    for issue in issues:
        line = f"- Delivery ID: {issue['delivery_id']}"

        if 'issue_type' in issue and issue['issue_type']:
            line += f" | Issue: {issue['issue_type']}"

        if 'severity' in issue and issue['severity']:
            line += f" | Severity: {issue['severity']}"

        lines.append(line)

        if include_details and issue.get('details'):
            lines.append(f"  Details: {issue['details']}")

    return '\n'.join(lines)


def format_issue_details(issues: List[Dict]) -> str:
    """
    Format detailed issue information.

    Args:
        issues: List of issue dicts

    Returns:
        Formatted string
    """
    lines = []

    for issue in issues:
        lines.append(f"Delivery: {issue['delivery_id']}")
        lines.append(f"  Type: {issue.get('issue_type', 'N/A')}")
        lines.append(f"  Severity: {issue.get('severity', 'N/A')}")
        lines.append(f"  Details: {issue.get('details', 'N/A')}")
        lines.append(f"  Expected: {issue.get('expected', 'N/A')}")
        lines.append(f"  Found: {issue.get('actual', 'N/A')}")
        lines.append("")

    return '\n'.join(lines)


def generate_email(
    template_name: str,
    contact_name: str,
    issues: List[Dict],
    **kwargs
) -> Dict[str, str]:
    """
    Generate email from template.

    Args:
        template_name: Template to use
        contact_name: Recipient name
        issues: List of issues
        **kwargs: Additional template variables

    Returns:
        Dict with subject and body
    """
    template = TEMPLATES.get(template_name, TEMPLATES['quality'])

    # Format issue lists
    issue_list = format_issue_list(issues)
    issue_details = format_issue_details(issues)

    # Prepare template variables
    variables = {
        'contact_name': contact_name,
        'count': len(issues),
        'issue_list': issue_list,
        'issue_details': issue_details,
        'date': datetime.now().strftime('%Y-%m-%d'),
        **kwargs
    }

    # Format template
    subject = template['subject'].format(**variables)
    body = template['body'].format(**variables)

    return {
        'subject': subject,
        'body': body
    }


def run_pod_email(
    issues_report: str,
    contacts_path: str = None,
    template: str = 'quality',
    output_folder: str = None,
    group_by: str = 'by-business'
) -> Path:
    """
    Main function to generate POD emails.

    Args:
        issues_report: Path to issues report
        contacts_path: Path to contacts Excel
        template: Template name
        output_folder: Path for email drafts
        group_by: Grouping mode

    Returns:
        Path to email log
    """
    issues_path = Path(issues_report)

    if not issues_path.exists():
        print(f"Error: Issues report not found: {issues_path}")
        sys.exit(1)

    print(f"Issues Report: {issues_path}")
    print(f"Template: {template}")
    print(f"Group By: {group_by}")
    print("-" * 50)

    # Load data
    print("Loading issues...")
    issues_df = load_issues_report(issues_path)
    print(f"Found {len(issues_df)} issues")

    # Load contacts
    contacts = {}
    if contacts_path:
        print(f"Loading contacts from: {contacts_path}")
        contacts = load_contacts(Path(contacts_path))
        print(f"Found {len(contacts)} contacts")

    # Group issues
    print("Grouping issues...")
    if group_by == 'by-business':
        grouped = group_issues_by_business(issues_df)
    else:
        grouped = group_issues_by_type(issues_df)

    print(f"Created {len(grouped)} groups")

    # Setup output
    config = PODConfig(output_folder=output_folder)
    drafts_folder = config.output_folder / 'email_drafts'
    drafts_folder.mkdir(parents=True, exist_ok=True)

    # Generate emails
    print("Generating emails...")
    email_log = []

    for group_name, issues in grouped.items():
        if not issues:
            continue

        # Get contact info
        contact_info = contacts.get(group_name, {})
        contact_name = contact_info.get('name', 'Team')
        contact_email = contact_info.get('email', '')

        # Generate email
        email = generate_email(
            template_name=template,
            contact_name=contact_name,
            issues=issues,
            total=len(issues_df),
            received=0,  # Would come from status report
            missing=0,
            issue_count=len(issues)
        )

        # Save draft
        safe_name = "".join(c for c in group_name if c.isalnum() or c in ' -_')[:30]
        draft_file = drafts_folder / f"email_{safe_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

        with open(draft_file, 'w', encoding='utf-8') as f:
            f.write(f"TO: {contact_email}\n")
            f.write(f"SUBJECT: {email['subject']}\n")
            f.write("-" * 50 + "\n\n")
            f.write(email['body'])

        email_log.append({
            'Group': group_name,
            'Contact Name': contact_name,
            'Contact Email': contact_email,
            'Subject': email['subject'],
            'Issue Count': len(issues),
            'Draft File': str(draft_file),
            'Generated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        })

        print(f"  Created: {draft_file.name}")

    # Create summary
    summary = {
        'Emails Generated': len(email_log),
        'Total Issues Covered': sum(e['Issue Count'] for e in email_log),
        'Template Used': template,
        'Generated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

    # Print summary
    print("-" * 50)
    print("SUMMARY")
    print("-" * 50)
    for key, value in summary.items():
        print(f"{key}: {value}")

    # Write log
    log_df = pd.DataFrame(email_log)
    log_path = config.get_output_path('pod_email_log')
    write_report(log_df, log_path, 'Email Log', summary)

    print("-" * 50)
    print(f"Email log saved: {log_path}")
    print(f"Drafts folder: {drafts_folder}")

    return log_path


def main():
    """Parse arguments and run POD email generation."""
    parser = argparse.ArgumentParser(
        description='POD Issue Email Generator'
    )
    parser.add_argument(
        'issues_report',
        help='Path to pod-issues report Excel'
    )
    parser.add_argument(
        '--contacts', '-c',
        help='Path to business contacts Excel'
    )
    parser.add_argument(
        '--template', '-t',
        choices=['missing', 'quality', 'resolution', 'summary'],
        default='quality',
        help='Email template to use'
    )
    parser.add_argument(
        '--group-by', '-g',
        choices=['by-business', 'by-type'],
        default='by-business',
        help='How to group issues'
    )
    parser.add_argument(
        '--output', '-o',
        help='Output folder for email drafts'
    )

    args = parser.parse_args()

    run_pod_email(
        issues_report=args.issues_report,
        contacts_path=args.contacts,
        template=args.template,
        output_folder=args.output,
        group_by=args.group_by
    )


if __name__ == '__main__':
    main()
