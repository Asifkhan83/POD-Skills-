"""
PDF utilities for POD skills.
OCR text extraction and field parsing for scanned POD documents.
"""
import re
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any

try:
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image
    # Set Tesseract and Poppler paths for Windows
    import os
    POPPLER_PATH = None
    if os.name == 'nt':  # Windows
        tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        if os.path.exists(tesseract_path):
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
        # Poppler path for pdf2image
        poppler_path = os.path.expanduser(r'~\poppler\poppler-24.08.0\Library\bin')
        if os.path.exists(poppler_path):
            POPPLER_PATH = poppler_path
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

try:
    from fuzzywuzzy import fuzz
    FUZZY_AVAILABLE = True
except ImportError:
    FUZZY_AVAILABLE = False


# Date patterns for extraction
DATE_PATTERNS = [
    # DD/MM/YYYY or MM/DD/YYYY
    (r'\b(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})\b', 'dmy_or_mdy'),
    # YYYY-MM-DD
    (r'\b(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})\b', 'ymd'),
    # DD/MM/YY or MM/DD/YY
    (r'\b(\d{1,2})[/\-](\d{1,2})[/\-](\d{2})\b', 'dmy_or_mdy_short'),
    # DD Month YYYY (15 January 2024)
    (r'\b(\d{1,2})\s+(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{4})\b', 'dMy'),
    # Month DD, YYYY (January 15, 2024)
    (r'\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{1,2}),?\s+(\d{4})\b', 'Mdy'),
]

MONTH_MAP = {
    'jan': 1, 'january': 1,
    'feb': 2, 'february': 2,
    'mar': 3, 'march': 3,
    'apr': 4, 'april': 4,
    'may': 5,
    'jun': 6, 'june': 6,
    'jul': 7, 'july': 7,
    'aug': 8, 'august': 8,
    'sep': 9, 'september': 9,
    'oct': 10, 'october': 10,
    'nov': 11, 'november': 11,
    'dec': 12, 'december': 12,
}


def extract_text_from_pdf(pdf_path: Path, use_ocr: bool = True) -> str:
    """
    Extract text from PDF file.

    Args:
        pdf_path: Path to PDF file
        use_ocr: If True, use OCR for scanned PDFs. If False, try text extraction first.

    Returns:
        Extracted text content
    """
    pdf_path = Path(pdf_path)

    if not pdf_path.exists():
        return ""

    text = ""

    # Try pdfplumber first (for text-based PDFs)
    if PDFPLUMBER_AVAILABLE and not use_ocr:
        try:
            import pdfplumber
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"

            # If we got substantial text, return it
            if len(text.strip()) > 50:
                return text.strip()
        except Exception:
            pass

    # Use OCR for scanned PDFs
    if OCR_AVAILABLE and use_ocr:
        try:
            # Convert PDF to images
            if POPPLER_PATH:
                images = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
            else:
                images = convert_from_path(pdf_path, dpi=300)

            for image in images:
                # Run OCR on each page
                page_text = pytesseract.image_to_string(image, lang='eng')
                text += page_text + "\n"

            return text.strip()
        except Exception as e:
            # Return error info for debugging
            return f"[OCR Error: {str(e)}]"

    # Fallback to pdfplumber if OCR not available
    if PDFPLUMBER_AVAILABLE:
        try:
            import pdfplumber
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
        except Exception:
            pass

    return text.strip()


def parse_dates_from_text(text: str) -> List[datetime]:
    """
    Extract all dates from text.

    Args:
        text: Text content to search

    Returns:
        List of parsed datetime objects
    """
    dates = []

    for pattern, format_type in DATE_PATTERNS:
        matches = re.finditer(pattern, text, re.IGNORECASE)

        for match in matches:
            try:
                if format_type == 'dmy_or_mdy':
                    # Assume DD/MM/YYYY for values where day > 12
                    a, b, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
                    if a > 12:
                        day, month = a, b
                    elif b > 12:
                        month, day = a, b
                    else:
                        # Ambiguous, assume DD/MM/YYYY
                        day, month = a, b
                    dates.append(datetime(year, month, day))

                elif format_type == 'ymd':
                    year, month, day = int(match.group(1)), int(match.group(2)), int(match.group(3))
                    dates.append(datetime(year, month, day))

                elif format_type == 'dmy_or_mdy_short':
                    a, b, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
                    year = 2000 + year if year < 100 else year
                    if a > 12:
                        day, month = a, b
                    elif b > 12:
                        month, day = a, b
                    else:
                        day, month = a, b
                    dates.append(datetime(year, month, day))

                elif format_type == 'dMy':
                    day = int(match.group(1))
                    month = MONTH_MAP.get(match.group(2).lower()[:3], 1)
                    year = int(match.group(3))
                    dates.append(datetime(year, month, day))

                elif format_type == 'Mdy':
                    month = MONTH_MAP.get(match.group(1).lower()[:3], 1)
                    day = int(match.group(2))
                    year = int(match.group(3))
                    dates.append(datetime(year, month, day))

            except (ValueError, IndexError):
                continue

    return dates


def parse_invoice_numbers_from_text(text: str) -> List[str]:
    """
    Extract potential invoice numbers from text.
    Looks for numeric sequences of 4-10 digits (typical invoice format).

    Args:
        text: Text content to search

    Returns:
        List of potential invoice numbers
    """
    # Look for invoice-related patterns
    invoice_patterns = [
        r'(?:invoice|inv|inv\.?\s*(?:no|number|#)?)[:\s#]*(\d{4,10})',  # Invoice: 12345
        r'(?:bill|receipt)[:\s#]*(\d{4,10})',  # Bill: 12345
        r'\b(\d{4,10})\b',  # Plain numbers 4-10 digits
    ]

    all_matches = []
    text_lower = text.lower()

    # First try invoice-specific patterns
    for pattern in invoice_patterns[:-1]:
        matches = re.findall(pattern, text_lower)
        all_matches.extend(matches)

    # If no invoice-specific matches, try plain numbers
    if not all_matches:
        matches = re.findall(invoice_patterns[-1], text)
        all_matches.extend(matches)

    # Return unique matches
    unique_matches = list(set(all_matches))
    unique_matches.sort(key=len, reverse=True)

    return unique_matches


def parse_delivery_ids_from_text(text: str) -> List[str]:
    """
    Extract potential delivery IDs from text.
    Looks for numeric sequences of 8-15 digits.

    Args:
        text: Text content to search

    Returns:
        List of potential delivery IDs
    """
    # Find numeric sequences (typically 10+ digits for delivery IDs)
    pattern = r'\b(\d{8,15})\b'
    matches = re.findall(pattern, text)

    # Return unique matches sorted by length (longer = more likely to be ID)
    unique_matches = list(set(matches))
    unique_matches.sort(key=len, reverse=True)

    return unique_matches


def parse_customer_names_from_text(text: str, known_customers: List[str] = None) -> List[str]:
    """
    Extract potential customer names from text.

    Args:
        text: Text content to search
        known_customers: Optional list of known customer names for matching

    Returns:
        List of potential customer name matches
    """
    found_names = []

    # If we have known customers, look for fuzzy matches
    if known_customers and FUZZY_AVAILABLE:
        # Get words from text
        words = text.split()

        for customer in known_customers:
            # Try matching 2-4 word combinations
            for n in range(2, 5):
                for i in range(len(words) - n + 1):
                    phrase = ' '.join(words[i:i+n])
                    ratio = fuzz.ratio(customer.lower(), phrase.lower())
                    if ratio >= 70:  # 70% threshold for extraction
                        found_names.append((customer, phrase, ratio))

    # Also look for common patterns like "Customer:", "Consignee:", etc.
    patterns = [
        r'(?:Customer|Consignee|Delivered to|Receiver|Ship to|Bill to)[:\s]+([A-Z][a-zA-Z\s&.,]+?)(?:\n|$)',
        r'(?:Name|Company)[:\s]+([A-Z][a-zA-Z\s&.,]+?)(?:\n|$)',
    ]

    for pattern in patterns:
        matches = re.findall(pattern, text)
        for match in matches:
            clean_match = match.strip().rstrip('.,')
            if len(clean_match) > 3:
                found_names.append(('_extracted_', clean_match, 100))

    return found_names


def parse_pod_fields(text: str, known_customers: List[str] = None) -> Dict[str, Any]:
    """
    Parse all available fields from POD text content.

    Args:
        text: Extracted text from POD
        known_customers: Optional list of known customer names

    Returns:
        Dictionary with extracted fields
    """
    fields = {
        'raw_text': text,
        'invoice_numbers': [],
        'delivery_ids': [],
        'dates': [],
        'customer_matches': [],
        'has_signature': False,
        'extracted_at': datetime.now().isoformat(),
    }

    if not text or text.startswith('[OCR Error'):
        fields['error'] = text if text else 'No text extracted'
        return fields

    # Extract invoice numbers (primary key)
    fields['invoice_numbers'] = parse_invoice_numbers_from_text(text)

    # Extract delivery IDs (fallback key)
    fields['delivery_ids'] = parse_delivery_ids_from_text(text)

    # Extract dates
    dates = parse_dates_from_text(text)
    fields['dates'] = [d.strftime('%Y-%m-%d') for d in dates]

    # Extract customer names
    customer_matches = parse_customer_names_from_text(text, known_customers)
    fields['customer_matches'] = customer_matches

    # Check for signature indicators
    signature_keywords = ['signature', 'signed', 'received by', 'receiver', 'sign here']
    fields['has_signature'] = any(kw in text.lower() for kw in signature_keywords)

    return fields


def compare_fields(
    pdf_fields: Dict[str, Any],
    manifest_row: Dict[str, Any],
    date_tolerance_days: int = 2,
    customer_match_threshold: int = 80
) -> Dict[str, Any]:
    """
    Compare extracted PDF fields against manifest row data.
    Uses Invoice Number as primary key, Delivery ID as fallback.

    Args:
        pdf_fields: Extracted fields from PDF
        manifest_row: Row data from manifest Excel
        date_tolerance_days: Allowed difference in days for date matching
        customer_match_threshold: Minimum fuzzy match score for customer

    Returns:
        Comparison results dictionary
    """
    results = {
        'invoice_match': False,
        'delivery_id_match': False,
        'id_match': False,  # Overall ID match (invoice or delivery)
        'date_match': False,
        'customer_match': False,
        'overall_match': 'No',
        'match_score': 0,
        'issues': [],
        'pdf_invoice': None,
        'pdf_delivery_id': None,
        'pdf_date': None,
        'pdf_customer': None,
        'manifest_invoice': str(manifest_row.get('invoice_number', '')),
        'manifest_delivery_id': str(manifest_row.get('delivery_id', '')),
        'manifest_date': str(manifest_row.get('date', '')),
        'manifest_customer': str(manifest_row.get('customer', '')),
    }

    total_checks = 3
    passed_checks = 0

    # Check Invoice Number (Primary Key)
    manifest_invoice = str(manifest_row.get('invoice_number', '')).strip()
    pdf_invoices = pdf_fields.get('invoice_numbers', [])

    if manifest_invoice and pdf_invoices:
        if manifest_invoice in pdf_invoices:
            results['invoice_match'] = True
            results['id_match'] = True
            results['pdf_invoice'] = manifest_invoice
            passed_checks += 1
        else:
            # Check for partial matches
            for pdf_inv in pdf_invoices:
                if manifest_invoice in pdf_inv or pdf_inv in manifest_invoice:
                    results['invoice_match'] = True
                    results['id_match'] = True
                    results['pdf_invoice'] = pdf_inv
                    passed_checks += 1
                    break

    # If invoice didn't match, try Delivery ID (Fallback Key)
    if not results['id_match']:
        manifest_id = str(manifest_row.get('delivery_id', '')).strip()
        pdf_ids = pdf_fields.get('delivery_ids', [])

        if manifest_id and pdf_ids:
            if manifest_id in pdf_ids:
                results['delivery_id_match'] = True
                results['id_match'] = True
                results['pdf_delivery_id'] = manifest_id
                passed_checks += 1
            else:
                # Check for partial matches
                for pdf_id in pdf_ids:
                    if manifest_id in pdf_id or pdf_id in manifest_id:
                        results['delivery_id_match'] = True
                        results['id_match'] = True
                        results['pdf_delivery_id'] = pdf_id
                        passed_checks += 1
                        break

    # Report ID mismatch if neither matched
    if not results['id_match']:
        results['pdf_invoice'] = pdf_invoices[0] if pdf_invoices else None
        results['pdf_delivery_id'] = pdf_fields.get('delivery_ids', [None])[0]

        if manifest_invoice:
            results['issues'].append(f"Invoice mismatch: PDF has {pdf_invoices[0] if pdf_invoices else 'none'}, expected {manifest_invoice}")
        elif manifest_row.get('delivery_id'):
            results['issues'].append(f"Delivery ID mismatch: PDF has {results['pdf_delivery_id'] or 'none'}, expected {manifest_row.get('delivery_id')}")

    # Check date
    manifest_date_str = str(manifest_row.get('date', ''))
    pdf_dates = pdf_fields.get('dates', [])

    if manifest_date_str and pdf_dates:
        try:
            # Parse manifest date
            manifest_date = None
            for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d %H:%M:%S']:
                try:
                    manifest_date = datetime.strptime(str(manifest_date_str).split()[0], fmt)
                    break
                except ValueError:
                    continue

            if manifest_date:
                for pdf_date_str in pdf_dates:
                    pdf_date = datetime.strptime(pdf_date_str, '%Y-%m-%d')
                    diff = abs((pdf_date - manifest_date).days)

                    if diff <= date_tolerance_days:
                        results['date_match'] = True
                        results['pdf_date'] = pdf_date_str
                        passed_checks += 1
                        break

                if not results['date_match']:
                    results['pdf_date'] = pdf_dates[0]
                    results['issues'].append(f"Date mismatch: PDF has {pdf_dates[0]}, expected {manifest_date.strftime('%Y-%m-%d')}")
        except Exception:
            results['issues'].append(f"Could not parse manifest date: {manifest_date_str}")
    elif not pdf_dates:
        results['issues'].append("No date found in PDF")

    # Check customer name
    manifest_customer = str(manifest_row.get('customer', '')).strip()
    customer_matches = pdf_fields.get('customer_matches', [])

    if manifest_customer and FUZZY_AVAILABLE:
        best_match = None
        best_score = 0

        for known, extracted, score in customer_matches:
            if known == manifest_customer and score > best_score:
                best_match = extracted
                best_score = score

        # Also check raw text for customer name
        raw_text = pdf_fields.get('raw_text', '')
        if raw_text:
            ratio = fuzz.partial_ratio(manifest_customer.lower(), raw_text.lower())
            if ratio > best_score:
                best_score = ratio
                best_match = f"Found in text (score: {ratio}%)"

        if best_score >= customer_match_threshold:
            results['customer_match'] = True
            results['pdf_customer'] = best_match
            passed_checks += 1
        else:
            results['pdf_customer'] = best_match if best_match else "Not found"
            results['issues'].append(f"Customer mismatch: best match score {best_score}%, expected {manifest_customer}")
    elif not manifest_customer:
        passed_checks += 1  # No customer to check
    else:
        results['issues'].append("Customer matching not available (fuzzywuzzy not installed)")

    # Calculate overall match
    results['match_score'] = int((passed_checks / total_checks) * 100)

    if passed_checks == total_checks:
        results['overall_match'] = 'Yes'
    elif passed_checks >= 2:
        results['overall_match'] = 'Partial'
    else:
        results['overall_match'] = 'No'

    return results


def check_ocr_available() -> Tuple[bool, str]:
    """
    Check if OCR dependencies are available.

    Returns:
        Tuple of (is_available, message)
    """
    if not OCR_AVAILABLE:
        return False, "OCR not available. Install: pip install pytesseract pdf2image Pillow"

    try:
        # Try to get Tesseract version
        version = pytesseract.get_tesseract_version()
        return True, f"Tesseract OCR v{version} available"
    except Exception as e:
        return False, f"Tesseract not found. Install from: https://github.com/UB-Mannheim/tesseract/wiki - Error: {e}"
