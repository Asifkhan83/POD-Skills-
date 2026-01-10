"""
Configuration management for POD skills.
Handles paths, column mappings, and settings.
"""
import os
from pathlib import Path
from datetime import datetime


class PODConfig:
    """Configuration for POD management skills."""

    # Default paths - override these in your environment
    DEFAULT_POD_FOLDER = r"D:\PODs"
    DEFAULT_MANIFEST_PATH = r"D:\Data\manifest.xlsx"
    DEFAULT_OUTPUT_FOLDER = r"D:\Reports"
    DEFAULT_ARCHIVE_FOLDER = r"D:\Archive"

    # Excel column mappings (adjust to match your manifest)
    MANIFEST_COLUMNS = {
        'delivery_id': 'Delivery ID',      # Column with delivery numbers
        'date': 'Delivery Date',            # Column with delivery date
        'customer': 'Customer Name',        # Column with customer name
        'status': 'Status',                 # Column with current status
    }

    # File patterns
    POD_FILE_EXTENSIONS = ['.pdf', '.PDF']

    # Validation settings
    DATE_TOLERANCE_DAYS = 2  # Allow +/- days for date matching
    CUSTOMER_MATCH_THRESHOLD = 80  # Fuzzy match percentage threshold

    def __init__(
        self,
        pod_folder: str = None,
        manifest_path: str = None,
        output_folder: str = None,
        archive_folder: str = None
    ):
        self.pod_folder = Path(pod_folder or self.DEFAULT_POD_FOLDER)
        self.manifest_path = Path(manifest_path) if manifest_path else Path(self.DEFAULT_MANIFEST_PATH)
        self.output_folder = Path(output_folder or self.DEFAULT_OUTPUT_FOLDER)
        self.archive_folder = Path(archive_folder or self.DEFAULT_ARCHIVE_FOLDER)

        # Ensure output folder exists
        self.output_folder.mkdir(parents=True, exist_ok=True)

    def get_output_path(self, prefix: str, extension: str = 'xlsx') -> Path:
        """Generate timestamped output file path."""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{prefix}_{timestamp}.{extension}"
        return self.output_folder / filename

    def validate_paths(self) -> dict:
        """Validate that required paths exist."""
        issues = {}

        if not self.pod_folder.exists():
            issues['pod_folder'] = f"POD folder not found: {self.pod_folder}"

        if not self.manifest_path.exists():
            issues['manifest_path'] = f"Manifest file not found: {self.manifest_path}"

        return issues

    @classmethod
    def from_args(cls, args: dict) -> 'PODConfig':
        """Create config from command line arguments."""
        return cls(
            pod_folder=args.get('pod_folder'),
            manifest_path=args.get('manifest'),
            output_folder=args.get('output'),
            archive_folder=args.get('archive')
        )


def parse_delivery_id(filename: str) -> str:
    """
    Extract delivery ID from filename.
    Handles formats like: 9354302576.pdf, DEL_9354302576.pdf
    """
    # Remove extension
    name = Path(filename).stem

    # Extract numeric part (handles various prefixes)
    import re
    numbers = re.findall(r'\d+', name)

    if numbers:
        # Return the longest numeric sequence (usually the delivery ID)
        return max(numbers, key=len)

    return name


def format_date(date_value) -> str:
    """Format date value consistently."""
    if isinstance(date_value, datetime):
        return date_value.strftime('%Y-%m-%d')
    if isinstance(date_value, str):
        return date_value
    return str(date_value) if date_value else ''
