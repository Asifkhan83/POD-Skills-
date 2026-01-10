"""
Create sample data for testing POD skills.
Generates:
- Manifest Excel with 20 deliveries
- 15 POD PDFs (5 missing, 3 extra)
"""
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import random

# Paths
SAMPLE_DIR = Path(__file__).parent
PODS_DIR = SAMPLE_DIR / "pods"
DATA_DIR = SAMPLE_DIR / "data"

# Sample customers
CUSTOMERS = [
    "ABC Logistics",
    "Metro Healthcare",
    "FastTrack Retail",
    "Global Pharma Inc",
    "City Hospital",
    "QuickMart Stores",
    "Prime Distributors",
    "MedSupply Co"
]

def create_manifest():
    """Create sample manifest Excel."""

    # Generate 20 delivery IDs
    delivery_ids = [
        "9354302576",
        "7104253522",
        "2641500014",
        "8825471039",
        "5563920187",
        "3347821456",
        "6219045783",
        "4478123690",
        "7891234567",
        "1234567890",
        "9876543210",
        "5555666677",
        "1112223334",
        "9998887776",
        "4443332221",
        "7776665554",
        "2223334445",
        "8889990001",
        "6667778889",
        "3334445556"
    ]

    # Generate dates (last 7 days)
    base_date = datetime.now()

    data = []
    for i, did in enumerate(delivery_ids):
        date = base_date - timedelta(days=random.randint(0, 6))
        customer = random.choice(CUSTOMERS)
        status = random.choice(["Delivered", "In Transit", "Pending"])

        data.append({
            "Delivery ID": did,
            "Delivery Date": date.strftime("%Y-%m-%d"),
            "Customer Name": customer,
            "Status": status
        })

    df = pd.DataFrame(data)
    manifest_path = DATA_DIR / "manifest.xlsx"
    df.to_excel(manifest_path, index=False)
    print(f"Created manifest: {manifest_path}")
    print(f"  - {len(df)} entries")

    return delivery_ids


def create_sample_pdfs(delivery_ids):
    """Create sample PDF files."""

    # Create 15 PDFs from manifest (leaving 5 missing)
    present_ids = delivery_ids[:15]

    # Create 3 extra PDFs not in manifest
    extra_ids = ["9999999999", "8888888888", "7777777777"]

    all_ids = present_ids + extra_ids

    for did in all_ids:
        pdf_path = PODS_DIR / f"{did}.pdf"

        # Create a minimal valid PDF
        pdf_content = b"""%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj
2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj
3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
>>
endobj
4 0 obj
<<
/Length 44
>>
stream
BT
/F1 12 Tf
100 700 Td
(POD Document - Delivery ID: """ + did.encode() + b""") Tj
ET
endstream
endobj
xref
0 5
0000000000 65535 f
0000000009 00000 n
0000000058 00000 n
0000000115 00000 n
0000000214 00000 n
trailer
<<
/Size 5
/Root 1 0 R
>>
startxref
308
%%EOF"""

        with open(pdf_path, 'wb') as f:
            f.write(pdf_content)

    print(f"Created {len(all_ids)} PDF files in: {PODS_DIR}")
    print(f"  - {len(present_ids)} matching manifest")
    print(f"  - {len(extra_ids)} extra (not in manifest)")
    print(f"  - {len(delivery_ids) - len(present_ids)} missing from manifest")

    return present_ids, extra_ids


def main():
    """Generate all sample data."""
    print("=" * 50)
    print("Creating Sample Data for POD Skills Testing")
    print("=" * 50)
    print()

    # Create manifest
    delivery_ids = create_manifest()
    print()

    # Create PDFs
    present_ids, extra_ids = create_sample_pdfs(delivery_ids)
    print()

    # Summary
    print("=" * 50)
    print("SAMPLE DATA SUMMARY")
    print("=" * 50)
    print(f"Manifest entries: 20")
    print(f"POD files created: 18")
    print(f"  - Present: 15")
    print(f"  - Missing: 5")
    print(f"  - Extra: 3")
    print()
    print("Missing POD IDs (for verification):")
    missing_ids = [did for did in delivery_ids if did not in present_ids]
    for did in missing_ids:
        print(f"  - {did}")
    print()
    print("Extra POD IDs (for verification):")
    for did in extra_ids:
        print(f"  - {did}")
    print()
    print("=" * 50)
    print("TEST COMMAND:")
    print("=" * 50)
    print(f'python pod-check/pod_check.py "{PODS_DIR}" "{DATA_DIR / "manifest.xlsx"}"')
    print()


if __name__ == "__main__":
    main()
