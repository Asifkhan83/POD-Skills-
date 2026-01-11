"""
POD Presence Validator - /pod-check
Compare scanned POD PDFs against manifest to find missing/extra PODs.

Usage:
    python pod_check.py [pod_folder] [manifest_path] [--output output_folder]
"""
import sys
import argparse
from pathlib import Path
from datetime import datetime
from typing import Set, Dict, List, Tuple

import pandas as pd

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))
from shared.config import PODConfig, parse_id_from_filename
from shared.excel_utils import read_manifest, write_report, apply_status_formatting
from shared.pdf_utils import extract_text_from_pdf, parse_pod_fields, compare_fields, check_ocr_available
from shared.report_utils import MarkdownReport, create_pod_check_report


def scan_pod_folder(folder: Path, extensions: List[str]) -> Tuple[Dict[str, Path], List[str]]:
    """
    Scan folder for POD files and extract IDs (invoice number or delivery ID from filename).

    Args:
        folder: Path to POD folder
        extensions: List of valid file extensions

    Returns:
        Tuple of (Dict mapping ID to file path, List of duplicate warnings)
    """
    pod_files = {}
    duplicates = []
    seen_files = set()  # Track files to avoid case-insensitive duplicates on Windows

    for ext in extensions:
        for file_path in folder.glob(f"*{ext}"):
            # Skip if we've already seen this file (case-insensitive)
            file_key = str(file_path).lower()
            if file_key in seen_files:
                continue
            seen_files.add(file_key)

            file_id = parse_id_from_filename(file_path.name)
            if file_id:
                if file_id in pod_files:
                    # Duplicate found - warn user (different files with same ID)
                    existing = pod_files[file_id]
                    if existing.name.lower() != file_path.name.lower():
                        duplicates.append(
                            f"Duplicate ID '{file_id}': '{file_path.name}' conflicts with '{existing.name}'"
                        )
                pod_files[file_id] = file_path

    return pod_files, duplicates


def compare_pods(
    manifest_ids: Set[str],
    scanned_ids: Set[str]
) -> Tuple[Set[str], Set[str], Set[str]]:
    """
    Compare manifest IDs with scanned POD IDs.

    Args:
        manifest_ids: Set of delivery IDs from manifest
        scanned_ids: Set of delivery IDs from scanned files

    Returns:
        Tuple of (present, missing, extra) sets
    """
    present = manifest_ids & scanned_ids
    missing = manifest_ids - scanned_ids
    extra = scanned_ids - manifest_ids

    return present, missing, extra


def extract_and_compare_content(
    pod_files: Dict[str, Path],
    manifest_df: pd.DataFrame,
    use_ocr: bool = True
) -> Dict[str, Dict]:
    """
    Extract content from PDFs and compare with manifest data.
    Uses Invoice Number as primary key, Delivery ID as fallback.

    Args:
        pod_files: Dict mapping ID to file path
        manifest_df: Manifest DataFrame with delivery details
        use_ocr: Whether to use OCR for text extraction

    Returns:
        Dict mapping ID to comparison results
    """
    results = {}

    # Get list of known customers for matching
    known_customers = []
    if 'customer' in manifest_df.columns:
        known_customers = manifest_df['customer'].dropna().unique().tolist()

    # Create lookup for manifest data (by invoice number and delivery ID)
    manifest_lookup = {}
    for _, row in manifest_df.iterrows():
        invoice_num = str(row.get('invoice_number', '')).strip()
        delivery_id = str(row.get('delivery_id', '')).strip()

        manifest_data = {
            'invoice_number': invoice_num,
            'delivery_id': delivery_id,
            'date': row.get('date', ''),
            'customer': row.get('customer', ''),
        }

        # Index by both invoice number and delivery ID
        if invoice_num and invoice_num != 'nan':
            manifest_lookup[invoice_num] = manifest_data
        if delivery_id and delivery_id != 'nan':
            manifest_lookup[delivery_id] = manifest_data

    # Process each PDF
    total = len(pod_files)
    for idx, (file_id, pdf_path) in enumerate(pod_files.items(), 1):
        print(f"  Processing {idx}/{total}: {pdf_path.name}...", end='\r')

        try:
            # Extract text from PDF
            text = extract_text_from_pdf(pdf_path, use_ocr=use_ocr)

            # Parse fields from text
            pdf_fields = parse_pod_fields(text, known_customers)

            # Compare with manifest if available
            manifest_row = manifest_lookup.get(file_id, {
                'invoice_number': file_id,
                'delivery_id': '',
                'date': '',
                'customer': '',
            })

            comparison = compare_fields(
                pdf_fields,
                manifest_row,
                date_tolerance_days=PODConfig.DATE_TOLERANCE_DAYS,
                customer_match_threshold=PODConfig.CUSTOMER_MATCH_THRESHOLD
            )

            results[file_id] = comparison

        except Exception as e:
            results[file_id] = {
                'overall_match': 'Error',
                'match_score': 0,
                'issues': [f"Processing error: {str(e)}"],
                'pdf_invoice': None,
                'pdf_delivery_id': None,
                'pdf_date': None,
                'pdf_customer': None,
            }

    print()  # New line after progress
    return results


def create_report_dataframe(
    manifest_df: pd.DataFrame,
    pod_files: Dict[str, Path],
    present: Set[str],
    missing: Set[str],
    extra: Set[str],
    comparison_results: Dict[str, Dict] = None
) -> pd.DataFrame:
    """
    Create report DataFrame with all POD statuses.
    Uses Invoice Number as primary identifier.

    Args:
        manifest_df: Original manifest DataFrame
        pod_files: Dict of ID to file path
        present: Set of present IDs
        missing: Set of missing IDs
        extra: Set of extra IDs
        comparison_results: Optional dict with content comparison results

    Returns:
        Report DataFrame
    """
    rows = []
    include_comparison = comparison_results is not None

    # Process manifest entries
    for _, row in manifest_df.iterrows():
        invoice_num = str(row.get('invoice_number', '')).strip()
        delivery_id = str(row.get('delivery_id', '')).strip()

        # Use invoice number as primary, delivery ID as fallback
        primary_id = invoice_num if invoice_num and invoice_num != 'nan' else delivery_id

        if not primary_id or primary_id == 'nan':
            continue

        if primary_id in present:
            status = 'Present'
            filename = pod_files.get(primary_id, Path()).name
        elif primary_id in missing:
            status = 'Missing'
            filename = ''
        else:
            continue

        row_data = {
            'Invoice Number': invoice_num if invoice_num != 'nan' else '',
            'Delivery ID': delivery_id if delivery_id != 'nan' else '',
            'Status': status,
            'Filename': filename,
            'Manifest Date': row.get('date', ''),
            'Customer': row.get('customer', ''),
        }

        # Add comparison data if available
        if include_comparison and primary_id in comparison_results:
            comp = comparison_results[primary_id]
            row_data.update({
                'Content Match': comp.get('overall_match', 'N/A'),
                'Match Score': f"{comp.get('match_score', 0)}%",
                'PDF Invoice': comp.get('pdf_invoice', ''),
                'Invoice Match': 'Yes' if comp.get('invoice_match') else 'No',
                'PDF Date': comp.get('pdf_date', ''),
                'Date Match': 'Yes' if comp.get('date_match') else 'No',
                'PDF Customer': comp.get('pdf_customer', ''),
                'Customer Match': 'Yes' if comp.get('customer_match') else 'No',
                'Issues': '; '.join(comp.get('issues', [])),
            })
        elif include_comparison:
            row_data.update({
                'Content Match': 'N/A' if status == 'Missing' else 'Not Processed',
                'Match Score': 'N/A',
                'PDF Invoice': '',
                'Invoice Match': 'N/A',
                'PDF Date': '',
                'Date Match': 'N/A',
                'PDF Customer': '',
                'Customer Match': 'N/A',
                'Issues': '',
            })

        rows.append(row_data)

    # Add extra PODs (not in manifest)
    for file_id in extra:
        file_path = pod_files.get(file_id, Path())
        row_data = {
            'Invoice Number': file_id,
            'Delivery ID': '',
            'Status': 'Extra',
            'Filename': file_path.name if file_path else '',
            'Manifest Date': 'N/A',
            'Customer': 'N/A (not in manifest)',
        }

        if include_comparison and file_id in comparison_results:
            comp = comparison_results[file_id]
            row_data.update({
                'Content Match': 'Extra (No Manifest)',
                'Match Score': 'N/A',
                'PDF Invoice': comp.get('pdf_invoice', ''),
                'Invoice Match': 'N/A',
                'PDF Date': comp.get('pdf_date', ''),
                'Date Match': 'N/A',
                'PDF Customer': comp.get('pdf_customer', ''),
                'Customer Match': 'N/A',
                'Issues': 'Not in manifest',
            })
        elif include_comparison:
            row_data.update({
                'Content Match': 'Extra (No Manifest)',
                'Match Score': 'N/A',
                'PDF Invoice': '',
                'Invoice Match': 'N/A',
                'PDF Date': '',
                'Date Match': 'N/A',
                'PDF Customer': '',
                'Customer Match': 'N/A',
                'Issues': 'Not in manifest',
            })

        rows.append(row_data)

    df = pd.DataFrame(rows)

    if not df.empty:
        # Sort: Missing first, then Present, then Extra
        status_order = {'Missing': 0, 'Present': 1, 'Extra': 2}
        df['_sort'] = df['Status'].map(status_order)
        df = df.sort_values(['_sort', 'Invoice Number']).drop('_sort', axis=1)

    return df


def run_pod_check(
    pod_folder: str = None,
    manifest_path: str = None,
    output_folder: str = None,
    compare_content: bool = False,
    use_ocr: bool = True,
    output_format: str = 'md'
) -> Path:
    """
    Main function to run POD presence check.

    Args:
        pod_folder: Path to folder with POD PDFs
        manifest_path: Path to manifest Excel
        output_folder: Path for output report
        compare_content: If True, extract and compare PDF content against manifest
        use_ocr: If True, use OCR for scanned PDFs (requires Tesseract)
        output_format: Output format - 'md', 'csv', 'xlsx', 'html', 'pdf', 'all'

    Returns:
        Path to generated report
    """
    # Initialize config
    config = PODConfig(
        pod_folder=pod_folder,
        manifest_path=manifest_path,
        output_folder=output_folder
    )

    # Validate paths
    issues = config.validate_paths()
    if issues:
        for key, msg in issues.items():
            print(f"Error: {msg}")
        sys.exit(1)

    print(f"POD Folder: {config.pod_folder}")
    print(f"Manifest: {config.manifest_path}")
    if compare_content:
        print(f"Content Comparison: Enabled (OCR: {'Yes' if use_ocr else 'No'})")
    print("-" * 50)

    # Check OCR availability if content comparison is enabled
    if compare_content and use_ocr:
        ocr_available, ocr_message = check_ocr_available()
        if not ocr_available:
            print(f"Warning: {ocr_message}")
            print("Falling back to text extraction without OCR...")
            use_ocr = False

    # Read manifest
    print("Reading manifest...")
    manifest_df = read_manifest(config.manifest_path, PODConfig.MANIFEST_COLUMNS)

    # Use Invoice Number as primary key, Delivery ID as fallback
    primary_key = PODConfig.PRIMARY_KEY
    fallback_key = PODConfig.FALLBACK_KEY

    # Build manifest IDs set (try primary key first, then fallback)
    manifest_ids = set()
    for _, row in manifest_df.iterrows():
        primary_val = str(row.get(primary_key, '')).strip()
        fallback_val = str(row.get(fallback_key, '')).strip()

        if primary_val and primary_val != 'nan':
            manifest_ids.add(primary_val)
        elif fallback_val and fallback_val != 'nan':
            manifest_ids.add(fallback_val)

    print(f"Found {len(manifest_ids)} entries in manifest (using {primary_key}, fallback: {fallback_key})")

    # Scan POD folder
    print("Scanning POD folder...")
    pod_files, duplicates = scan_pod_folder(config.pod_folder, PODConfig.POD_FILE_EXTENSIONS)
    scanned_ids = set(pod_files.keys())
    print(f"Found {len(scanned_ids)} POD files")

    # Warn about duplicates
    if duplicates:
        print(f"\nWARNING: Found {len(duplicates)} duplicate delivery ID(s):")
        for dup in duplicates:
            print(f"  - {dup}")
        print("  (Only the last file for each ID will be used)\n")

    # Compare presence
    print("Comparing presence...")
    present, missing, extra = compare_pods(manifest_ids, scanned_ids)

    # Content comparison (optional)
    comparison_results = None
    if compare_content and present:
        print(f"Extracting and comparing PDF content ({len(present)} files)...")
        # Only compare present PODs (those that exist and are in manifest)
        present_files = {did: pod_files[did] for did in present if did in pod_files}
        comparison_results = extract_and_compare_content(
            present_files, manifest_df, use_ocr=use_ocr
        )

    # Create report
    report_df = create_report_dataframe(
        manifest_df, pod_files, present, missing, extra,
        comparison_results=comparison_results
    )

    # Summary statistics
    summary = {
        'Total in Manifest': len(manifest_ids),
        'PODs Present': len(present),
        'PODs Missing': len(missing),
        'Extra PODs': len(extra),
        'Match Rate': f"{len(present) / len(manifest_ids) * 100:.1f}%" if manifest_ids else "N/A",
    }

    # Add content comparison stats if enabled
    if comparison_results:
        full_match = sum(1 for r in comparison_results.values() if r.get('overall_match') == 'Yes')
        partial_match = sum(1 for r in comparison_results.values() if r.get('overall_match') == 'Partial')
        no_match = sum(1 for r in comparison_results.values() if r.get('overall_match') == 'No')
        errors = sum(1 for r in comparison_results.values() if r.get('overall_match') == 'Error')

        summary.update({
            'Content Comparison': 'Enabled',
            'Full Match': full_match,
            'Partial Match': partial_match,
            'No Match': no_match,
            'Errors': errors,
            'Content Match Rate': f"{(full_match + partial_match) / len(comparison_results) * 100:.1f}%" if comparison_results else "N/A",
        })

    summary['Generated'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Print summary
    print("-" * 50)
    print("SUMMARY")
    print("-" * 50)
    for key, value in summary.items():
        print(f"{key}: {value}")

    # Create report
    report = create_pod_check_report(summary, report_df, "POD Check Report")

    # Save report in specified format(s)
    base_path = config.get_output_path('pod_check_report')
    print("-" * 50)

    if output_format.lower() == 'all':
        # Save in all formats
        saved_files = report.save_all(base_path.with_suffix(''))
        print("Reports saved:")
        for fmt, path in saved_files.items():
            print(f"  [{fmt.upper()}] {path}")
        output_path = saved_files.get('md', base_path)
    else:
        # Save in single format
        output_path = report.save(base_path, output_format)
        print(f"Report saved: {output_path}")

    # Also print Markdown to console for immediate viewing
    if output_format.lower() in ['md', 'markdown', 'all']:
        print("-" * 50)
        print("\n" + report.to_markdown())

    return output_path


def main():
    """Parse arguments and run POD check."""
    parser = argparse.ArgumentParser(
        description='POD Presence Validator - Compare scanned PODs against manifest'
    )
    parser.add_argument(
        'pod_folder',
        nargs='?',
        help='Path to folder containing POD PDFs'
    )
    parser.add_argument(
        'manifest_path',
        nargs='?',
        help='Path to manifest Excel file'
    )
    parser.add_argument(
        '--output', '-o',
        help='Output folder for report'
    )
    parser.add_argument(
        '--compare-content', '-c',
        action='store_true',
        help='Extract PDF content and compare against manifest data'
    )
    parser.add_argument(
        '--no-ocr',
        action='store_true',
        help='Disable OCR (use text extraction only, for text-based PDFs)'
    )
    parser.add_argument(
        '--format', '-f',
        default='md',
        choices=['md', 'csv', 'xlsx', 'html', 'pdf', 'all'],
        help='Output format: md (Markdown), csv, xlsx (Excel), html, pdf, or all'
    )

    args = parser.parse_args()

    run_pod_check(
        pod_folder=args.pod_folder,
        manifest_path=args.manifest_path,
        output_folder=args.output,
        compare_content=args.compare_content,
        use_ocr=not args.no_ocr,
        output_format=args.format
    )


if __name__ == '__main__':
    main()
