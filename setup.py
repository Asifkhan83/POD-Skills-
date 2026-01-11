"""
POD Skills Setup - Run once to create required folders and verify installation.

Usage:
    python setup.py
"""
import subprocess
import sys
from pathlib import Path


def create_folders():
    """Create required folders."""
    folders = [
        r"D:\PODs",
        r"D:\Data",
        r"D:\Reports",
        r"D:\Archive",
    ]

    print("Creating required folders...")
    for folder in folders:
        path = Path(folder)
        if not path.exists():
            path.mkdir(parents=True, exist_ok=True)
            print(f"  Created: {folder}")
        else:
            print(f"  Exists:  {folder}")


def check_dependencies():
    """Check if required packages are installed."""
    required = ['pandas', 'openpyxl', 'pdfplumber', 'fuzzywuzzy']
    missing = []

    print("\nChecking dependencies...")
    for package in required:
        try:
            __import__(package)
            print(f"  OK: {package}")
        except ImportError:
            print(f"  MISSING: {package}")
            missing.append(package)

    return missing


def install_dependencies(packages):
    """Install missing packages."""
    print(f"\nInstalling missing packages: {', '.join(packages)}")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install'] + packages)


def create_sample_contacts():
    """Create sample contacts file."""
    import pandas as pd

    contacts_path = Path(r"D:\Data\contacts.xlsx")

    if not contacts_path.exists():
        print("\nCreating sample contacts file...")
        df = pd.DataFrame([
            {'Business Name': 'Sample Business', 'Contact Email': 'email@example.com', 'Contact Name': 'Contact Person'},
        ])
        df.to_excel(contacts_path, index=False)
        print(f"  Created: {contacts_path}")
        print("  NOTE: Edit this file with your actual business contacts")
    else:
        print(f"\nContacts file exists: {contacts_path}")


def main():
    """Run setup."""
    print("=" * 60)
    print("POD Skills Setup")
    print("=" * 60)

    # Create folders
    create_folders()

    # Check dependencies
    missing = check_dependencies()
    if missing:
        install_dependencies(missing)
        print("\nDependencies installed successfully!")

    # Create sample contacts
    try:
        create_sample_contacts()
    except Exception as e:
        print(f"\nNote: Could not create contacts file: {e}")

    print("\n" + "=" * 60)
    print("Setup Complete!")
    print("=" * 60)
    print("""
Next steps:
1. Copy your manifest Excel file to: D:\\Data\\manifest.xlsx
2. Edit D:\\Data\\contacts.xlsx with your business contacts
3. Ensure POD PDFs are syncing to: D:\\PODs
4. Run: python daily_workflow.py

Or double-click: run_daily_workflow.bat
""")


if __name__ == '__main__':
    main()
