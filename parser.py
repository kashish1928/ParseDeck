"""
ParseDeck - PDF Text Extraction Engine
Specialized for multi-column academic and database research papers.
Preserves natural reading order across two-column layouts using PyMuPDF (fitz).
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple

try:
    import pymupdf as fitz  # PyMuPDF
except ImportError:
    import fitz  # Fallback for legacy installations

logger = logging.getLogger("parsedeck.parser")


@dataclass
class TextBlock:
    """Represents an extracted text block with coordinate metadata."""

    x0: float
    y0: float
    x1: float
    y1: float
    text: str
    block_no: int
    column_idx: int = 0

    @property
    def width(self) -> float:
        return self.x1 - self.x0

    @property
    def height(self) -> float:
        return self.y1 - self.y0

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2.0


@dataclass
class ParsedPage:
    """Represents a single extracted page with structured blocks and merged text."""

    page_number: int  # 1-indexed
    width: float
    height: float
    blocks: List[TextBlock] = field(default_factory=list)
    text: str = ""


@dataclass
class ParsedDocument:
    """Represents the complete parsed academic paper."""

    file_path: Path
    total_pages: int
    parsed_pages: List[ParsedPage] = field(default_factory=list)
    full_text: str = ""
    metadata: dict = field(default_factory=dict)

    def get_preview(self, max_chars: int = 1000) -> str:
        """Returns a trimmed preview of the parsed document text."""
        return self.full_text[:max_chars] + ("..." if len(self.full_text) > max_chars else "")


def clean_academic_text(raw_text: str) -> str:
    """
    Cleans raw text extracted from PDF:
    - Re-joins hyphenated words across linebreaks (e.g., 'trans-\\naction' -> 'transaction').
    - Normalizes excessive whitespace and newlines while preserving paragraph breaks.
    - Strips non-printable control characters.
    """
    if not raw_text:
        return ""

    # Fix hyphenation across linebreaks: word- \n word -> wordword
    text = re.sub(r"(\b[A-Za-z]+)-\s*\n\s*([A-Za-z]+\b)", r"\1\2", raw_text)

    # Normalize carriage returns and tabs
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\t", " ")

    # Collapse repeated whitespace within lines
    lines = [re.sub(r"[ ]{2,}", " ", line.strip()) for line in text.split("\n")]

    # Reassemble paragraphs: collapse more than two consecutive newlines
    cleaned = "\n".join(lines)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned.strip()


def sort_blocks_two_column(
    blocks: List[Tuple[float, float, float, float, str, int, int]],
    page_width: float,
    page_height: float,
    column_split_threshold: Optional[float] = None,
) -> List[TextBlock]:
    """
    Sorts PyMuPDF blocks by column and vertical coordinate to preserve reading order.

    Academic papers typically follow:
    1. Full-width header (Paper Title, Authors, Abstract, Keywords) spanning across both columns.
    2. Left column (x0 < midpoint) sorted vertically by y0.
    3. Right column (x0 >= midpoint) sorted vertically by y0.
    4. Full-width footer (Spanning figures, tables, or footnotes) at the bottom.

    Args:
        blocks: Raw blocks from page.get_text("blocks")
                Each block is (x0, y0, x1, y1, text, block_no, block_type)
        page_width: Width of the PDF page in points.
        page_height: Height of the PDF page in points.
        column_split_threshold: Horizontal boundary splitting columns (defaults to page_width / 2).

    Returns:
        List of TextBlock objects ordered by natural reading flow.
    """
    mid_x = column_split_threshold if column_split_threshold is not None else (page_width / 2.0)

    # Filter out image blocks (block_type != 0) and empty strings
    text_blocks_raw = [b for b in blocks if len(b) >= 7 and b[6] == 0 and b[4].strip()]

    if not text_blocks_raw:
        return []

    # Identify full-width blocks: width > 65% of page width spanning across mid_x
    full_width_blocks: List[TextBlock] = []
    left_column_blocks: List[TextBlock] = []
    right_column_blocks: List[TextBlock] = []

    for b in text_blocks_raw:
        x0, y0, x1, y1, text, block_no, _ = b
        cleaned_block_text = clean_academic_text(text)
        if not cleaned_block_text:
            continue

        block_width = x1 - x0
        is_spanning = (block_width > 0.65 * page_width) and (x0 < mid_x * 0.9 and x1 > mid_x * 1.1)

        if is_spanning:
            full_width_blocks.append(
                TextBlock(x0=x0, y0=y0, x1=x1, y1=y1, text=cleaned_block_text, block_no=block_no, column_idx=-1)
            )
        else:
            # Multi-column split by x0 coordinate
            # If block starts before mid_x, it belongs to left column; else right column
            if x0 < mid_x:
                left_column_blocks.append(
                    TextBlock(x0=x0, y0=y0, x1=x1, y1=y1, text=cleaned_block_text, block_no=block_no, column_idx=0)
                )
            else:
                right_column_blocks.append(
                    TextBlock(x0=x0, y0=y0, x1=x1, y1=y1, text=cleaned_block_text, block_no=block_no, column_idx=1)
                )

    # Sort each column partition vertically by y0
    left_column_blocks.sort(key=lambda b: b.y0)
    right_column_blocks.sort(key=lambda b: b.y0)
    full_width_blocks.sort(key=lambda b: b.y0)

    # If no spanning blocks exist or no multi-column split detected, simple 2-column or 1-column assembly
    if not full_width_blocks:
        # Standard two-column order: left column (top to bottom) -> right column (top to bottom)
        return left_column_blocks + right_column_blocks

    # Determine vertical boundary where columns start
    # Full-width blocks above the top of the two columns are headers (title/abstract)
    min_col_y = min(
        [b.y0 for b in left_column_blocks + right_column_blocks]
    ) if (left_column_blocks or right_column_blocks) else page_height

    max_col_y = max(
        [b.y1 for b in left_column_blocks + right_column_blocks]
    ) if (left_column_blocks or right_column_blocks) else 0.0

    top_spanning = [b for b in full_width_blocks if b.y0 < min_col_y]
    bottom_spanning = [b for b in full_width_blocks if b.y0 >= max_col_y]
    mid_spanning = [b for b in full_width_blocks if b not in top_spanning and b not in bottom_spanning]

    ordered_blocks: List[TextBlock] = []
    ordered_blocks.extend(top_spanning)
    ordered_blocks.extend(left_column_blocks)
    ordered_blocks.extend(mid_spanning)
    ordered_blocks.extend(right_column_blocks)
    ordered_blocks.extend(bottom_spanning)

    return ordered_blocks


def extract_page(page: fitz.Page, page_number: int) -> ParsedPage:
    """Extracts and orders text blocks from a single PyMuPDF page."""
    rect = page.rect
    raw_blocks = page.get_text("blocks")

    sorted_blocks = sort_blocks_two_column(
        blocks=raw_blocks,
        page_width=rect.width,
        page_height=rect.height,
    )

    page_text = "\n\n".join([b.text for b in sorted_blocks])

    return ParsedPage(
        page_number=page_number,
        width=rect.width,
        height=rect.height,
        blocks=sorted_blocks,
        text=page_text,
    )


def parse_pdf(pdf_path: str | Path, max_pages: Optional[int] = None) -> ParsedDocument:
    """
    Parses an academic PDF into a structured document with preserved two-column reading order.

    Args:
        pdf_path: Path to the research paper PDF.
        max_pages: Optional maximum number of pages to parse (defaults to all pages).

    Returns:
        ParsedDocument containing page-level blocks, metadata, and full extracted text.

    Raises:
        FileNotFoundError: If the PDF path does not exist.
        ValueError: If the PDF cannot be opened or parsed.
    """
    path = Path(pdf_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"PDF file not found at: {path}")

    logger.info(f"Opening PDF document: {path.name}")

    try:
        doc = fitz.open(str(path))
    except Exception as exc:
        raise ValueError(f"Failed to open PDF with PyMuPDF: {exc}") from exc

    total_pages = len(doc)
    pages_to_process = min(total_pages, max_pages) if max_pages else total_pages
    logger.info(f"Parsing {pages_to_process} of {total_pages} pages...")

    parsed_pages: List[ParsedPage] = []
    combined_texts: List[str] = []

    for i in range(pages_to_process):
        page = doc[i]
        parsed_page = extract_page(page, page_number=i + 1)
        parsed_pages.append(parsed_page)

        if parsed_page.text:
            combined_texts.append(f"--- PAGE {i + 1} ---\n{parsed_page.text}")

    full_text = "\n\n".join(combined_texts)
    doc_metadata = doc.metadata or {}

    doc.close()

    logger.info(
        f"Extracted {len(parsed_pages)} pages ({len(full_text.split())} words) successfully."
    )

    return ParsedDocument(
        file_path=path,
        total_pages=total_pages,
        parsed_pages=parsed_pages,
        full_text=full_text,
        metadata=doc_metadata,
    )


if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    if len(sys.argv) < 2:
        print("Usage: python parser.py <path_to_pdf> [max_pages]")
        sys.exit(1)

    input_pdf = sys.argv[1]
    limit = int(sys.argv[2]) if len(sys.argv) > 2 else None

    try:
        parsed_doc = parse_pdf(input_pdf, max_pages=limit)
        print(f"\nSuccessfully parsed: {parsed_doc.file_path.name}")
        print(f"Total Pages in file: {parsed_doc.total_pages}")
        print(f"Parsed Pages: {len(parsed_doc.parsed_pages)}")
        print("\n--- Preview of First 500 Characters ---")
        print(parsed_doc.get_preview(500))
    except Exception as err:
        print(f"Error parsing PDF: {err}", file=sys.stderr)
        sys.exit(1)
