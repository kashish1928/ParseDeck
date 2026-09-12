"""
Tests for ParseDeck pipeline components:
- Schema validation in extractor.py
- Two-column PDF parsing in parser.py
- Slide deck generation in generator.py
- CLI interface in main.py
"""

import json
import subprocess
import sys
import unittest
from pathlib import Path
try:
    import pymupdf as fitz
except ImportError:
    import fitz

from extractor import AlgorithmStep, BenchmarkResult, KeyComponent, PaperAnalysis
from generator import generate_presentation
from parser import parse_pdf, sort_blocks_two_column


class TestExtractorSchema(unittest.TestCase):
    def test_schema_instantiation(self):
        analysis = PaperAnalysis(
            title="L-Store: A Hybrid OLTP/OLAP Engine with Hardware-Accelerated Vectorization",
            authors=["Alice Smith", "Bob Jones", "Carol White"],
            publication_venue="VLDB 2024",
            executive_summary="L-Store introduces a dual-format column store that merges row updates into vectorized columnar blocks with zero-copy SIMD processing.",
            core_problem="Contemporary analytical engines suffer from severe memory stall cycles when converting transactional row updates into vectorized column layouts.",
            problem_limitations=[
                "Row-to-column conversion creates CPU cache thrashing during high-velocity streaming writes.",
                "Traditional lock managers induce severe lock contention under concurrent OLAP scans.",
                "Compaction stages incur 35% background I/O amplification.",
            ],
            key_innovation="A lock-free lineage-indexed columnar log with hardware-directed SIMD merging that avoids copy stalls.",
            key_components=[
                KeyComponent(name="Lineage Index", description="Lock-free trie mapping logical record IDs to delta columnar blocks."),
                KeyComponent(name="SIMD Merger", description="AVX-512 accelerated vector processing for zero-copy delta unification."),
                KeyComponent(name="Async Compactor", description="Log-structured background vacuum engine with rate-limiting."),
            ],
            algorithm_overview="Queries scan immutable baseline columnar chunks and concurrently probe the lineage index to merge active updates on-the-fly using vectorized SIMD registers.",
            algorithm_steps=[
                AlgorithmStep(step_number=1, title="Ingestion & WAL Log", description="Transactions append uncompressed row records to the NVMe-backed write-ahead log."),
                AlgorithmStep(step_number=2, title="Lineage Trie Tagging", description="The lineage index atomically tags records with monotonic version timestamps."),
                AlgorithmStep(step_number=3, title="Vectorized SIMD Scan", description="OLAP query operators stream columnar vectors directly into CPU SIMD registers without row reassembly."),
            ],
            benchmark_results=[
                BenchmarkResult(
                    benchmark_name="TPC-H SF100",
                    baseline="PostgreSQL 16 with Citus",
                    result_metric="4.8x higher query throughput",
                    significance="Demonstrates dramatic acceleration on analytical aggregation queries.",
                ),
                BenchmarkResult(
                    benchmark_name="Hybrid CH-benCHmark",
                    baseline="DuckDB v0.9 (Concurrent Updates)",
                    result_metric="62% reduction in P99 latency",
                    significance="Proves that lock-free lineage indexing eliminates scan stalls during heavy writes.",
                ),
            ],
            system_tradeoffs_and_implications="Requires AVX-512 or ARM Neon support for optimal vector merging; consumes ~12% higher RAM for lineage pointer tables.",
            conclusion_takeaways=[
                "Zero-copy vectorized delta merging eliminates the boundary between transactional logs and columnar stores.",
                "Hardware-accelerated SIMD instructions can replace heavyweight write-lock synchronization.",
                "Applicable to modern real-time streaming analytics engines.",
            ],
        )

        self.assertEqual(len(analysis.benchmark_results), 2)
        self.assertEqual(analysis.publication_venue, "VLDB 2024")
        json_data = analysis.model_dump_json()
        restored = PaperAnalysis.model_validate_json(json_data)
        self.assertEqual(restored.title, analysis.title)


class TestParserTwoColumn(unittest.TestCase):
    def setUp(self):
        self.test_dir = Path("./scratch_test_files")
        self.test_dir.mkdir(parents=True, exist_ok=True)
        self.sample_pdf = self.test_dir / "sample_two_column_paper.pdf"

        # Create a synthetic 2-column academic PDF using PyMuPDF
        doc = fitz.open()
        page = doc.new_page(width=600, height=800)  # Standard letter-like size

        # 1. Full-width Title & Abstract at top
        page.insert_text(fitz.Point(100, 50), "Paper Title: Fast Columnar Database Engines", fontsize=16)
        page.insert_text(fitz.Point(100, 80), "Authors: A. Researcher, B. Engineer", fontsize=11)
        page.insert_textbox(
            fitz.Rect(50, 110, 550, 180),
            "Abstract: This paper presents an ultra-fast in-memory database engine that accelerates queries.",
            fontsize=10,
        )

        # 2. Left Column (x: 50 -> 280)
        page.insert_textbox(
            fitz.Rect(50, 200, 280, 700),
            "1. Introduction\nModern analytical query processing requires high memory bandwidth.\n"
            "We propose a novel index structure that minimizes cache misses.",
            fontsize=10,
        )

        # 3. Right Column (x: 320 -> 550)
        page.insert_textbox(
            fitz.Rect(320, 200, 550, 700),
            "2. System Architecture\nThe architecture consists of three modular components:\n"
            "the ingestion buffer, the columnar delta storage, and the vectorized query runner.",
            fontsize=10,
        )

        doc.save(str(self.sample_pdf))
        doc.close()

    def tearDown(self):
        if self.sample_pdf.exists():
            self.sample_pdf.unlink()
        if self.test_dir.exists():
            for f in self.test_dir.glob("*"):
                f.unlink()
            self.test_dir.rmdir()

    def test_parse_pdf_reading_order(self):
        parsed = parse_pdf(self.sample_pdf)
        self.assertEqual(parsed.total_pages, 1)
        self.assertIn("Paper Title", parsed.full_text)
        self.assertIn("1. Introduction", parsed.full_text)
        self.assertIn("2. System Architecture", parsed.full_text)

        # Ensure Introduction appears before System Architecture in the extracted text!
        intro_pos = parsed.full_text.find("1. Introduction")
        arch_pos = parsed.full_text.find("2. System Architecture")
        self.assertTrue(intro_pos != -1)
        self.assertTrue(arch_pos != -1)
        self.assertLess(intro_pos, arch_pos, "Left column must be extracted before right column!")


class TestPresentationGenerator(unittest.TestCase):
    def setUp(self):
        self.test_dir = Path("./scratch_test_files")
        self.test_dir.mkdir(parents=True, exist_ok=True)
        self.out_pptx = self.test_dir / "test_presentation.pptx"

    def tearDown(self):
        if self.out_pptx.exists():
            self.out_pptx.unlink()
        if self.test_dir.exists():
            for f in self.test_dir.glob("*"):
                f.unlink()
            self.test_dir.rmdir()

    def test_generate_presentation(self):
        sample_analysis = PaperAnalysis(
            title="Accelerating Database Analytics with Hardware Disaggregation",
            authors=["Jane Doe", "John Smith"],
            publication_venue="SIGMOD 2024",
            executive_summary="This paper investigates RDMA-based memory pooling for distributed database analytics.",
            core_problem="Network serialization overhead dominates distributed join processing in disaggregated storage.",
            problem_limitations=[
                "TCP stack overhead creates high CPU utilization.",
                "Shuffle phases stall query execution pipes.",
            ],
            key_innovation="One-sided RDMA zero-copy shuffle buffers with hardware kernel bypass.",
            key_components=[
                KeyComponent(name="RDMA Shuffle Ring", description="Circular ring buffers mapped directly into NIC memory."),
                KeyComponent(name="Vectorized Hash Join", description="SIMD partitioned hash join operator."),
            ],
            algorithm_overview="Compute nodes stream partition keys directly to remote memory via InfiniBand RDMA write operations without CPU intervention.",
            algorithm_steps=[
                AlgorithmStep(step_number=1, title="Hash Partitioning", description="Local SIMD radix partitioning of relations."),
                AlgorithmStep(step_number=2, title="One-Sided RDMA Write", description="Direct remote memory DMA transfer without RPC serialization."),
            ],
            benchmark_results=[
                BenchmarkResult(
                    benchmark_name="TPC-H SF1000",
                    baseline="Spark SQL with TCP Shuffle",
                    result_metric="5.4x speedup on Query 17",
                    significance="Eliminates CPU shuffle serialization bottleneck.",
                )
            ],
            system_tradeoffs_and_implications="Requires RoCEv2 or InfiniBand network fabrics; sensitive to packet drop retries.",
            conclusion_takeaways=[
                "Disaggregated memory architectures redefine query execution.",
                "Zero-copy networking is mandatory for 100Gbps+ fabrics.",
            ],
        )

        path = generate_presentation(sample_analysis, output_path=self.out_pptx)
        self.assertTrue(path.exists())
        self.assertGreater(path.stat().st_size, 1000)

        # Inspect generated presentation
        from pptx import Presentation
        pptx_doc = Presentation(str(path))
        self.assertEqual(len(pptx_doc.slides), 6, "Expected 6 slides generated.")


class TestCLI(unittest.TestCase):
    def test_cli_help(self):
        result = subprocess.run(
            [sys.executable, "main.py", "--help"],
            capture_output=True,
            text=True,
        )
        self.assertEqual(result.returncode, 0)
        self.assertIn("ParseDeck", result.stdout)
        self.assertIn("--output", result.stdout)
        self.assertIn("--model", result.stdout)


if __name__ == "__main__":
    unittest.main()
