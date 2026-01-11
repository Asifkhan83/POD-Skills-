"""
POD Daily Workflow - Run all skills in sequence.

Usage:
    python daily_workflow.py [manifest_path]
    python daily_workflow.py                          # Use default paths
    python daily_workflow.py "D:\Data\manifest.xlsx"  # Specify manifest

This script runs:
1. pod-check   - Find missing/extra PODs
2. pod-issues  - Detect quality issues
3. pod-status  - Consolidate status
4. pod-email   - Generate issue emails
5. pod-archive - Archive completed PODs (optional)
"""
import sys
import subprocess
from pathlib import Path
from datetime import datetime

# Configuration
SKILLS_DIR = Path(__file__).parent
POD_FOLDER = r"D:\PODs"
MANIFEST_PATH = r"D:\Data\manifest.xlsx"
OUTPUT_FOLDER = r"D:\Reports"
ARCHIVE_FOLDER = r"D:\Archive"
CONTACTS_PATH = r"D:\Data\contacts.xlsx"  # Optional


def run_skill(skill_name: str, args: list) -> bool:
    """Run a skill and return success status."""
    script_path = SKILLS_DIR / skill_name / f"{skill_name.replace('-', '_')}.py"

    if not script_path.exists():
        print(f"  ERROR: Script not found: {script_path}")
        return False

    cmd = [sys.executable, str(script_path)] + args

    try:
        result = subprocess.run(cmd, capture_output=False, text=True)
        return result.returncode == 0
    except Exception as e:
        print(f"  ERROR: {e}")
        return False


def main():
    """Run the complete daily workflow."""
    print("=" * 60)
    print("POD DAILY WORKFLOW")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # Get manifest path from args or use default
    manifest = sys.argv[1] if len(sys.argv) > 1 else MANIFEST_PATH

    # Validate paths
    if not Path(POD_FOLDER).exists():
        print(f"ERROR: POD folder not found: {POD_FOLDER}")
        print("Please update POD_FOLDER in this script.")
        sys.exit(1)

    if not Path(manifest).exists():
        print(f"ERROR: Manifest not found: {manifest}")
        print("Please provide manifest path as argument or update MANIFEST_PATH.")
        sys.exit(1)

    # Ensure output folder exists
    Path(OUTPUT_FOLDER).mkdir(parents=True, exist_ok=True)

    print(f"\nPOD Folder:  {POD_FOLDER}")
    print(f"Manifest:    {manifest}")
    print(f"Output:      {OUTPUT_FOLDER}")

    results = {}

    # Step 1: POD Check
    print("\n" + "-" * 60)
    print("STEP 1: Running /pod-check (Presence Validation)")
    print("-" * 60)
    results['pod-check'] = run_skill('pod-check', [
        POD_FOLDER, manifest, '--output', OUTPUT_FOLDER
    ])

    # Step 2: POD Issues
    print("\n" + "-" * 60)
    print("STEP 2: Running /pod-issues (Quality Check)")
    print("-" * 60)
    results['pod-issues'] = run_skill('pod-issues', [
        POD_FOLDER, manifest, '--output', OUTPUT_FOLDER
    ])

    # Find the latest reports for pod-status
    reports_dir = Path(OUTPUT_FOLDER)
    check_reports = sorted(reports_dir.glob('pod_check_report_*.xlsx'))
    issues_reports = sorted(reports_dir.glob('pod_issues_report_*.xlsx'))

    check_report = str(check_reports[-1]) if check_reports else ""
    issues_report = str(issues_reports[-1]) if issues_reports else ""

    # Step 3: POD Status
    print("\n" + "-" * 60)
    print("STEP 3: Running /pod-status (Status Consolidation)")
    print("-" * 60)
    status_args = [manifest, '--output', OUTPUT_FOLDER]
    if check_report:
        status_args.extend(['--check-report', check_report])
    if issues_report:
        status_args.extend(['--issues-report', issues_report])

    results['pod-status'] = run_skill('pod-status', status_args)

    # Step 4: POD Email
    print("\n" + "-" * 60)
    print("STEP 4: Running /pod-email (Email Generation)")
    print("-" * 60)
    if issues_report:
        email_args = [issues_report, '--template', 'quality', '--output', OUTPUT_FOLDER]
        if Path(CONTACTS_PATH).exists():
            email_args.extend(['--contacts', CONTACTS_PATH])
        results['pod-email'] = run_skill('pod-email', email_args)
    else:
        print("  SKIPPED: No issues report found")
        results['pod-email'] = None

    # Summary
    print("\n" + "=" * 60)
    print("WORKFLOW COMPLETE")
    print("=" * 60)
    print(f"Finished: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    print("Results:")
    for skill, success in results.items():
        status = "OK" if success else ("SKIPPED" if success is None else "FAILED")
        print(f"  {skill}: {status}")

    print(f"\nReports saved to: {OUTPUT_FOLDER}")

    # List generated reports
    print("\nGenerated Reports:")
    today = datetime.now().strftime('%Y%m%d')
    for report in sorted(reports_dir.glob(f'*_{today}_*.xlsx')):
        print(f"  - {report.name}")

    drafts_dir = reports_dir / 'email_drafts'
    if drafts_dir.exists():
        drafts = list(drafts_dir.glob(f'*_{today}_*.txt'))
        if drafts:
            print(f"\nEmail Drafts ({len(drafts)} files):")
            print(f"  - {drafts_dir}")


if __name__ == '__main__':
    main()
