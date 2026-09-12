#!/usr/bin/env python3
"""
ParseDeck - Academic Research Paper to PowerPoint Deck Generator
Command-line interface orchestrating the Parser -> Extractor -> Generator pipeline.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
from pathlib import Path

from dotenv import load_dotenv

from extractor import OpenRouterExtractor, PaperAnalysis
from generator import generate_presentation
from parser import parse_pdf

# Load environment variables (.env)
load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("parsedeck")


BANNER = r"""
  ____                      ____             _    
 |  _ \ __ _ _ __ ___  ___ |  _ \  ___  ___| | __
 | |_) / _` | '__/ __|/ _ \| | | |/ _ \/ __| |/ /
 |  __/ (_| | |  \__ \  __/| |_| |  __/ (__|   < 
 |_|   \__,_|_|  |___/\___||____/ \___|\___|_|\_\
      Academic Paper to Presentation Deck Engine
"""


def build_arg_parser() -> argparse.ArgumentParser:
    """Constructs the command-line argument parser."""
    parser = argparse.ArgumentParser(
        prog="parsedeck",
        description="ParseDeck: Ingest academic database research papers (PDFs) and generate executive PowerPoint decks.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  parsedeck paper.pdf -o presentation.pptx
  parsedeck paper.pdf --model anthropic/claude-3.5-sonnet --max-pages 12
  parsedeck paper.pdf -o vldb_deck.pptx --save-json paper_meta.json
        """,
    )

    parser.add_argument(
        "pdf",
        type=str,
        help="Path to the research paper PDF file to ingest.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        default=None,
        help="Path for the output .pptx file (default: <paper_stem>_presentation.pptx).",
    )
    parser.add_argument(
        "-m",
        "--model",
        type=str,
        default=os.getenv("OPENROUTER_MODEL", "google/gemini-2.5-flash"),
        help="OpenRouter model ID to use for extraction (default: google/gemini-2.5-flash or OPENROUTER_MODEL).",
    )
    parser.add_argument(
        "-k",
        "--api-key",
        type=str,
        default=None,
        help="OpenRouter API key (defaults to OPENROUTER_API_KEY environment variable).",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=None,
        help="Limit number of pages to parse from the PDF (useful for skipping appendix or citations).",
    )
    parser.add_argument(
        "--save-json",
        type=str,
        default=None,
        help="Optional path to save the intermediate structured JSON metadata extracted by OpenRouter.",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable detailed debug logging.",
    )

    return parser


def run_pipeline(
    pdf_path: str,
    output_path: str | None = None,
    model: str | None = None,
    api_key: str | None = None,
    max_pages: int | None = None,
    save_json_path: str | None = None,
) -> Path:
    """
    Orchestrates the 3-stage ParseDeck pipeline:
      1. Parser: Multi-column PDF text extraction using PyMuPDF.
      2. Extractor: Structured academic extraction via OpenRouter and Pydantic.
      3. Generator: Deck compilation with python-pptx.

    Returns:
        Path to the generated .pptx presentation file.
    """
    input_file = Path(pdf_path).resolve()
    if not input_file.is_file():
        raise FileNotFoundError(f"Target PDF file does not exist: {input_file}")

    if not output_path:
        output_file = input_file.parent / f"{input_file.stem}_presentation.pptx"
    else:
        output_file = Path(output_path).resolve()

    start_time = time.time()

    # -------------------------------------------------------------------------
    # Stage 1: PyMuPDF Multi-Column Extraction
    # -------------------------------------------------------------------------
    logger.info("[1/3] Ingesting and parsing PDF layout...")
    parsed_doc = parse_pdf(pdf_path=input_file, max_pages=max_pages)
    logger.info(
        f"      Extracted {len(parsed_doc.parsed_pages)} pages "
        f"({len(parsed_doc.full_text.split())} words) with reading order preserved."
    )

    if not parsed_doc.full_text.strip():
        raise ValueError("No extractable text found in the provided PDF.")

    # -------------------------------------------------------------------------
    # Stage 2: OpenRouter LLM Extraction & Validation
    # -------------------------------------------------------------------------
    logger.info(f"[2/3] Querying OpenRouter for technical metadata (Model: {model})...")
    extractor = OpenRouterExtractor(api_key=api_key, model=model)
    analysis: PaperAnalysis = extractor.extract(paper_text=parsed_doc.full_text)

    logger.info(f"      Identified Paper: '{analysis.title}'")
    if analysis.publication_venue:
        logger.info(f"      Venue: {analysis.publication_venue}")
    logger.info(f"      Extracted {len(analysis.benchmark_results)} empirical benchmark results.")
    logger.info(f"      Extracted {len(analysis.key_components)} core architectural components.")

    # Optionally persist structured JSON metadata
    if save_json_path:
        json_dest = Path(save_json_path).resolve()
        json_dest.parent.mkdir(parents=True, exist_ok=True)
        json_dest.write_text(
            analysis.model_dump_json(indent=2),
            encoding="utf-8",
        )
        logger.info(f"      Saved structured JSON analysis to: {json_dest}")

    # -------------------------------------------------------------------------
    # Stage 3: python-pptx Presentation Generation
    # -------------------------------------------------------------------------
    logger.info(f"[3/3] Compiling PowerPoint deck...")
    final_deck_path = generate_presentation(analysis=analysis, output_path=output_file)

    elapsed = time.time() - start_time
    file_size_kb = final_deck_path.stat().st_size / 1024.0
    logger.info(
        f"Done in {elapsed:.2f}s! Generated presentation saved to:\n      -> {final_deck_path} "
        f"({file_size_kb:.1f} KB, 6 slides)"
    )

    return final_deck_path


def main():
    """CLI entry point."""
    parser = build_arg_parser()
    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    print(BANNER)

    try:
        run_pipeline(
            pdf_path=args.pdf,
            output_path=args.output,
            model=args.model,
            api_key=args.api_key,
            max_pages=args.max_pages,
            save_json_path=args.save_json,
        )
    except KeyboardInterrupt:
        print("\nAborted by user.")
        sys.exit(130)
    except Exception as err:
        logger.error(f"Execution failed: {err}")
        if args.verbose:
            logger.exception("Detailed stack trace:")
        sys.exit(1)


if __name__ == "__main__":
    main()
