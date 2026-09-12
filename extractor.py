"""
ParseDeck - OpenRouter Academic Extraction Engine
Integrates with OpenRouter API using Pydantic for structured JSON validation.
Extracts: Core Problem, Key Innovation, Algorithm Overview, and Benchmark Results.
"""

from __future__ import annotations

import json
import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests
from dotenv import load_dotenv
from pydantic import BaseModel, Field, ValidationError

# Load environment variables from .env if present
load_dotenv()

logger = logging.getLogger("parsedeck.extractor")

DEFAULT_OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "google/gemma-4-31b-it:free")
OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"


# ==============================================================================
# Pydantic Schemas for Structured Paper Analysis
# ==============================================================================

class BenchmarkResult(BaseModel):
    """Structured empirical evaluation metric and baseline comparison."""

    benchmark_name: str = Field(
        ...,
        description="Name of the benchmark, dataset, or workload (e.g., 'TPC-H SF100', 'YCSB Workload A', 'ClickBench').",
    )
    baseline: str = Field(
        ...,
        description="The competing baseline system or previous state-of-the-art (e.g., 'PostgreSQL 16', 'DuckDB v0.9', 'Vanilla RocksDB').",
    )
    result_metric: str = Field(
        ...,
        description="The quantitative performance metric achieved (e.g., '3.8x throughput speedup', '42% lower P99 latency', '2.5x memory reduction').",
    )
    significance: str = Field(
        ...,
        description="Brief technical context on why this result demonstrates the paper's advantage.",
    )


class AlgorithmStep(BaseModel):
    """A distinct phase or step in the proposed algorithm or execution flow."""

    step_number: int = Field(..., description="1-indexed sequence order.")
    title: str = Field(..., description="Short descriptive title of the phase (e.g., 'Zero-Copy Vector Ingestion').")
    description: str = Field(..., description="Concise explanation of the operations performed in this step.")


class KeyComponent(BaseModel):
    """Core architectural component or data structure introduced in the system."""

    name: str = Field(..., description="Name of the component or subsystem (e.g., 'Adaptive Trie Index', 'Async Buffer Manager').")
    description: str = Field(..., description="Technical role and inner mechanics of this component.")


class PaperAnalysis(BaseModel):
    """Complete structured technical breakdown of an academic research paper."""

    title: str = Field(..., description="Full paper title.")
    authors: List[str] = Field(default_factory=list, description="Key authors or primary research group.")
    publication_venue: Optional[str] = Field(None, description="Conference or journal venue (e.g., VLDB, SIGMOD, OSDI, SOSP, arXiv).")
    executive_summary: str = Field(
        ...,
        description="High-level 2-3 sentence technical briefing of the paper's core hypothesis, solution, and outcome.",
    )
    core_problem: str = Field(
        ...,
        description="The primary architectural bottleneck, theoretical limitation, or performance problem being addressed.",
    )
    problem_limitations: List[str] = Field(
        ...,
        description="3 to 4 specific architectural bottlenecks or failure modes of existing state-of-the-art approaches.",
    )
    key_innovation: str = Field(
        ...,
        description="The primary technical breakthrough, novel paradigm shift, or core architectural mechanism introduced.",
    )
    key_components: List[KeyComponent] = Field(
        ...,
        description="3 to 4 modular architectural building blocks or data structures that realize the innovation.",
    )
    algorithm_overview: str = Field(
        ...,
        description="High-level description of the end-to-end algorithmic flow or query execution pipeline.",
    )
    algorithm_steps: List[AlgorithmStep] = Field(
        ...,
        description="3 to 5 chronological execution stages or algorithm phases detailing the mechanics.",
    )
    benchmark_results: List[BenchmarkResult] = Field(
        ...,
        description="Key empirical evaluation findings showing concrete comparisons against baselines.",
    )
    system_tradeoffs_and_implications: str = Field(
        ...,
        description="Real-world system implications: hardware requirements, trade-offs, memory/compute overheads, or ideal workload profiles.",
    )
    conclusion_takeaways: List[str] = Field(
        ...,
        description="3 bullet points highlighting why this research matters and its long-term impact on systems engineering.",
    )


# ==============================================================================
# OpenRouter Extraction Client
# ==============================================================================

class OpenRouterExtractor:
    """Client for extracting structured research contributions via OpenRouter."""

    def __init__(
        self,
        api_key: Optional[str] = None,
        model: Optional[str] = None,
        timeout: int = 120,
    ):
        """
        Initializes the OpenRouter extractor.

        Args:
            api_key: OpenRouter API key. If not provided, reads OPENROUTER_API_KEY environment variable.
            model: Model identifier on OpenRouter. Defaults to OPENROUTER_MODEL env var or google/gemini-2.5-flash.
            timeout: HTTP request timeout in seconds.
        """
        self.api_key = api_key or os.getenv("OPENROUTER_API_KEY")
        if not self.api_key:
            raise ValueError(
                "OpenRouter API key is required. Set the OPENROUTER_API_KEY environment variable "
                "or pass api_key to OpenRouterExtractor."
            )

        self.model = model or DEFAULT_OPENROUTER_MODEL
        self.timeout = timeout

    def _build_system_prompt(self) -> str:
        """Constructs the system prompt enforcing rigorous academic systems extraction."""
        schema_json = json.dumps(PaperAnalysis.model_json_schema(), indent=2)

        return (
            "You are an elite principal systems researcher and database architect specializing in analyzing "
            "premier database and systems papers (e.g., VLDB, SIGMOD, OSDI, SOSP, EuroSys, CIDR).\n\n"
            "Your task is to extract the core technical contributions, architectural designs, algorithms, "
            "and quantitative benchmark results from the provided research paper text into a strictly validated JSON structure.\n\n"
            "CRITICAL EXTRACTION GUIDELINES:\n"
            "1. Be precise and technically deep. Avoid vague marketing buzzwords; use exact architectural terms "
            "(e.g., 'lock-free Bw-tree', 'SIMD vectorized execution', 'NVMe-oF latency stalls', 'paged attention').\n"
            "2. Extract concrete empirical benchmarks. Identify the workloads (e.g., TPC-C, YCSB, ClickBench), "
            "the baseline systems (e.g., PostgreSQL, RocksDB, DuckDB), and exact speedups/reductions.\n"
            "3. Structure the algorithm overview into sequential, logical phases.\n"
            "4. Output ONLY valid JSON adhering strictly to the following Pydantic JSON schema without any surrounding text.\n\n"
            f"SCHEMA:\n{schema_json}"
        )

    def _clean_json_response(self, content: str) -> str:
        """Extracts JSON substring if the LLM wraps it in markdown code fences."""
        content = content.strip()
        # Look for ```json ... ``` blocks
        json_fence_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", content)
        if json_fence_match:
            return json_fence_match.group(1).strip()
        return content

    def extract(self, paper_text: str, max_chars: int = 120_000) -> PaperAnalysis:
        """
        Extracts structured paper contributions from the extracted paper text.

        Args:
            paper_text: Full or truncated text extracted from the academic PDF.
            max_chars: Character limit to prevent excessive token consumption.

        Returns:
            Validated PaperAnalysis Pydantic instance.

        Raises:
            RuntimeError: If OpenRouter API returns an error or fails validation.
        """
        if not paper_text.strip():
            raise ValueError("Input paper text is empty. Cannot extract metadata.")

        trimmed_text = paper_text[:max_chars]
        if len(paper_text) > max_chars:
            logger.warning(
                f"Paper text truncated from {len(paper_text)} to {max_chars} characters to fit context window."
            )

        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "HTTP-Referer": "https://github.com/kashish1928/ParseDeck",
            "X-Title": "ParseDeck Research Paper Analyzer",
            "Content-Type": "application/json",
        }

        user_prompt = (
            f"Analyze the following academic systems research paper and extract its core technical metadata:\n\n"
            f"{trimmed_text}\n\n"
            f"Respond with a complete, valid JSON object conforming to the PaperAnalysis schema."
        )

        payload: Dict[str, Any] = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": self._build_system_prompt()},
                {"role": "user", "content": user_prompt},
            ],
            "response_format": {"type": "json_object"},
            "temperature": 0.1,
        }

        logger.info(f"Sending extraction request to OpenRouter (Model: {self.model})...")

        try:
            response = requests.post(
                OPENROUTER_API_URL,
                headers=headers,
                json=payload,
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise RuntimeError(f"Network error communicating with OpenRouter: {exc}") from exc

        if response.status_code != 200:
            error_detail = response.text
            try:
                err_json = response.json()
                error_detail = err_json.get("error", {}).get("message", response.text)
            except Exception:
                pass
            raise RuntimeError(
                f"OpenRouter API returned HTTP {response.status_code}: {error_detail}"
            )

        resp_data = response.json()
        try:
            raw_content = resp_data["choices"][0]["message"]["content"]
        except (KeyError, IndexError) as exc:
            raise RuntimeError(f"Malformed response format from OpenRouter: {resp_data}") from exc

        cleaned_json = self._clean_json_response(raw_content)

        logger.info("Validating OpenRouter response with Pydantic schema...")
        try:
            analysis = PaperAnalysis.model_validate_json(cleaned_json)
            logger.info(f"Successfully validated paper metadata for: '{analysis.title}'")
            return analysis
        except ValidationError as val_err:
            logger.error(f"Pydantic schema validation failed. Raw payload:\n{cleaned_json[:500]}...")
            raise RuntimeError(f"Failed to validate OpenRouter output against PaperAnalysis schema: {val_err}") from val_err


def extract_contributions(
    paper_text: str,
    api_key: Optional[str] = None,
    model: Optional[str] = None,
) -> PaperAnalysis:
    """Convenience functional wrapper for extracting paper metadata."""
    extractor = OpenRouterExtractor(api_key=api_key, model=model)
    return extractor.extract(paper_text=paper_text)


if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    if len(sys.argv) < 2:
        print("Usage: python extractor.py <path_to_text_file>")
        sys.exit(1)

    input_file = Path(sys.argv[1])
    if not input_file.is_file():
        print(f"File not found: {input_file}", file=sys.stderr)
        sys.exit(1)

    text = input_file.read_text(encoding="utf-8")
    try:
        res = extract_contributions(text)
        print(f"\nTitle: {res.title}")
        print(f"Executive Summary: {res.executive_summary}\n")
        print("Core Problem:")
        print(f"  {res.core_problem}")
        print("\nKey Innovation:")
        print(f"  {res.key_innovation}")
        print("\nBenchmarks Extracted:")
        for bm in res.benchmark_results:
            print(f"  - [{bm.benchmark_name}] vs {bm.baseline}: {bm.result_metric} ({bm.significance})")
    except Exception as e:
        print(f"Extraction failed: {e}", file=sys.stderr)
        sys.exit(1)
