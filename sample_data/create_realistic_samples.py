"""
Create more realistic sample PDFs for testing POD issues detection.
Some PDFs will have correct data, some will have issues.
"""
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
import random

SAMPLE_DIR = Path(__file__).parent
PODS_DIR = SAMPLE_DIR / "pods"
DATA_DIR = SAMPLE_DIR / "data"

def create_pdf_with_content(filepath: Path, delivery_id: str, date: str, customer: str):
    """Create a PDF with actual text content."""

    # PDF with embedded text content
    content = f"""PROOF OF DELIVERY

Delivery ID: {delivery_id}
Date: {date}
Customer: {customer}

Signature: [SIGNED]
Driver: John Smith

This document confirms delivery completion."""

    # Create PDF with text
    pdf_content = f"""%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>
endobj
4 0 obj
<< /Length {len(content) + 50} >>
stream
BT
/F1 12 Tf
50 700 Td
({content.replace(chr(10), ') Tj 0 -15 Td (')}) Tj
ET
endstream
endobj
5 0 obj
<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>
endobj
xref
0 6
0000000000 65535 f
0000000009 00000 n
0000000058 00000 n
0000000115 00000 n
0000000266 00000 n
trailer
<< /Size 6 /Root 1 0 R >>
startxref
400
%%EOF"""

    with open(filepath, 'w') as f:
        f.write(pdf_content)


def main():
    """Create realistic sample data."""
    print("Creating realistic sample PDFs...")

    # Read manifest
    manifest_path = DATA_DIR / "manifest.xlsx"
    df = pd.read_excel(manifest_path)

    # Clear existing PDFs
    for pdf in PODS_DIR.glob("*.pdf"):
        pdf.unlink()

    created = 0
    issues_created = {
        'correct': 0,
        'date_wrong': 0,
        'customer_wrong': 0,
        'no_date': 0
    }

    for idx, row in df.iterrows():
        delivery_id = str(row['Delivery ID'])
        manifest_date = row['Delivery Date']
        manifest_customer = row['Customer Name']

        # Skip some to create "missing" PODs
        if idx >= 15:
            continue

        pdf_path = PODS_DIR / f"{delivery_id}.pdf"

        # Create different scenarios
        scenario = idx % 5

        if scenario == 0:
            # Correct data
            create_pdf_with_content(pdf_path, delivery_id, manifest_date, manifest_customer)
            issues_created['correct'] += 1

        elif scenario == 1:
            # Wrong date (off by 5 days)
            wrong_date = (datetime.strptime(manifest_date, '%Y-%m-%d') + timedelta(days=5)).strftime('%Y-%m-%d')
            create_pdf_with_content(pdf_path, delivery_id, wrong_date, manifest_customer)
            issues_created['date_wrong'] += 1

        elif scenario == 2:
            # Wrong customer
            wrong_customer = "ACME Corporation"
            create_pdf_with_content(pdf_path, delivery_id, manifest_date, wrong_customer)
            issues_created['customer_wrong'] += 1

        elif scenario == 3:
            # No date in PDF
            create_pdf_with_content(pdf_path, delivery_id, "N/A", manifest_customer)
            issues_created['no_date'] += 1

        else:
            # Correct data
            create_pdf_with_content(pdf_path, delivery_id, manifest_date, manifest_customer)
            issues_created['correct'] += 1

        created += 1

    # Add extra PDFs not in manifest
    for extra_id in ["9999999999", "8888888888", "7777777777"]:
        pdf_path = PODS_DIR / f"{extra_id}.pdf"
        create_pdf_with_content(pdf_path, extra_id, "2026-01-10", "Unknown Customer")
        created += 1

    print(f"\nCreated {created} PDFs:")
    print(f"  - Correct data: {issues_created['correct']}")
    print(f"  - Wrong date: {issues_created['date_wrong']}")
    print(f"  - Wrong customer: {issues_created['customer_wrong']}")
    print(f"  - No date: {issues_created['no_date']}")
    print(f"  - Extra (not in manifest): 3")
    print(f"  - Missing (no PDF): 5")

    print("\nReady for testing!")


if __name__ == "__main__":
    main()
