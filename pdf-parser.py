"""pdf-parser.py

Stream, search and extract structured data from very large PDF files (page-by-page)
to avoid excessive memory use. Writes incremental CSV output with one row per match.

Usage examples:
  python pdf-parser.py --input report.pdf --output matches.csv
  python pdf-parser.py --input "*.pdf" --output matches.csv --patterns patterns.json

Patterns JSON format (key -> regex string):
  {
	"account_number": "Acct(?:ount)?\\s*(?:No\\.?|#)?\\s*[:\\-]?\\s*([A-Za-z0-9\\-]{4,40})",
	"taxpayer_name": "(?:Taxpayer|Owner|Name)\\s*[:\\-]\\s*([A-Z][A-Za-z\\'\\.\\s,]{3,100})"
  }

This script uses PyMuPDF (installed as "PyMuPDF") for fast page text extraction.
"""

from __future__ import annotations

import argparse
import csv
import glob
import json
import logging
import os
import re
import sys
import time
from typing import Dict, Iterator, List, Optional, Pattern, Tuple

import fitz  # PyMuPDF

DEFAULT_PATTERNS: Dict[str, str] = {
	# Generic account-looking pattern (very permissive). Users should supply their
	# own more specific patterns via --patterns for best results.
	"account_number": r"\bAcct(?:ount)?\s*(?:No\.?|#)?\s*[:\-]?\s*([A-Za-z0-9\-]{4,64})\b",
	# Taxpayer/owner name heuristic (captures capitalized name-like strings after common labels)
	"taxpayer_name": r"(?:Taxpayer|Owner|Name)\s*[:\-]\s*([A-Z][A-Za-z\'\.\s,\-]{2,120})",
}


def compile_patterns(patterns: Dict[str, str]) -> Dict[str, Pattern]:
	compiled: Dict[str, Pattern] = {}
	for k, v in patterns.items():
		try:
			compiled[k] = re.compile(v)
		except re.error as e:
			raise ValueError(f"Invalid regex for pattern '{k}': {e}")
	return compiled


def pdf_page_texts(path: str, password: Optional[str] = None) -> Iterator[Tuple[int, str]]:
	"""Yield (page_number (1-based), text) for each page in the PDF.

	Uses PyMuPDF which is efficient for large documents. If the PDF is encrypted and
	a password is provided, it attempts to open it.
	"""
	doc = fitz.open(path)
	if doc.needs_pass and password:
		try:
			doc.authenticate(password)
		except Exception:
			logging.warning("Failed to authenticate PDF with provided password")
	for i, page in enumerate(doc, start=1):
		try:
			text = page.get_text("text")
		except Exception as e:
			logging.exception("Failed to extract text from page %s of %s: %s", i, path, e)
			text = ""
		yield i, text


def find_matches_on_page(text: str, compiled: Dict[str, Pattern], context: int = 40) -> List[Dict]:
	"""Return a list of match dicts found in the given text for all compiled patterns."""
	results: List[Dict] = []
	for name, pat in compiled.items():
		for m in pat.finditer(text):
			span = m.span()
			start, end = span
			snippet = text[max(0, start - context) : min(len(text), end + context)].replace("\n", " ")
			# If the regex uses capture groups, prefer the first group's text; otherwise full match
			matched_text = m.group(1) if len(m.groups()) >= 1 else m.group(0)
			results.append({
				"pattern": name,
				"match": matched_text.strip(),
				"start": start,
				"end": end,
				"snippet": snippet.strip(),
			})
	return results


def process_file(
	input_path: str,
	compiled: Dict[str, Pattern],
	writer: csv.DictWriter,
	password: Optional[str] = None,
	progress_interval: int = 1000,
) -> int:
	"""Process a single PDF and write matches via the provided CSV writer.
	Returns the number of matches written.
	"""
	matches_written = 0
	start_time = time.time()
	try:
		for page_no, text in pdf_page_texts(input_path, password=password):
			for m in find_matches_on_page(text, compiled):
				row = {
					"file": os.path.basename(input_path),
					"path": os.path.abspath(input_path),
					"page": page_no,
					"pattern": m["pattern"],
					"match": m["match"],
					"start": m["start"],
					"end": m["end"],
					"snippet": m["snippet"],
				}
				writer.writerow(row)
				matches_written += 1

			if page_no % progress_interval == 0:
				elapsed = time.time() - start_time
				logging.info(
					"Processed %s pages of %s in %.1fs (matches so far: %d)",
					page_no,
					input_path,
					elapsed,
					matches_written,
				)
	except Exception:
		logging.exception("Error processing file %s", input_path)
	return matches_written


def expand_inputs(pattern: str) -> List[str]:
	# Accept a single file, wildcard, or directory
	if os.path.isdir(pattern):
		# find PDFs in directory
		return sorted(glob.glob(os.path.join(pattern, "*.pdf")))
	if any(ch in pattern for ch in "*?["):
		return sorted(glob.glob(pattern))
	return [pattern]


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
	p = argparse.ArgumentParser(description="Stream and extract patterns from large PDF files")
	p.add_argument("--input", "-i", required=True, help="Input PDF, glob (e.g. '*.pdf') or directory")
	p.add_argument("--output", "-o", required=True, help="Output CSV file (will append if exists)")
	p.add_argument(
		"--patterns",
		"-p",
		help="JSON file with name->regex mapping. If omitted, built-in heuristics are used.",
	)
	p.add_argument("--password", help="Password for encrypted PDFs (if needed)")
	p.add_argument("--context", type=int, default=40, help="Characters of context to save around matches")
	p.add_argument("--progress-interval", type=int, default=1000, help="Log progress every N pages")
	p.add_argument("--append", action="store_true", help="Append to existing CSV instead of overwriting")
	p.add_argument("--verbose", "-v", action="store_true", help="Verbose logging")
	return p.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
	args = parse_args(argv)
	logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

	# Load patterns
	patterns = DEFAULT_PATTERNS.copy()
	if args.patterns:
		with open(args.patterns, "r", encoding="utf-8") as pf:
			user = json.load(pf)
			if not isinstance(user, dict):
				logging.error("Patterns file must contain a JSON object mapping names to regex strings")
				return 2
			patterns.update(user)

	try:
		compiled = compile_patterns(patterns)
	except ValueError as e:
		logging.error(str(e))
		return 3

	inputs = expand_inputs(args.input)
	if not inputs:
		logging.error("No input PDFs found for: %s", args.input)
		return 4

	write_header = not (args.append and os.path.exists(args.output))
	out_dir = os.path.dirname(os.path.abspath(args.output))
	if out_dir and not os.path.exists(out_dir):
		os.makedirs(out_dir, exist_ok=True)

	fieldnames = ["file", "path", "page", "pattern", "match", "start", "end", "snippet"]
	mode = "a" if args.append else "w"
	total_matches = 0
	with open(args.output, mode, newline="", encoding="utf-8-sig") as csvfile:
		writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
		if write_header:
			writer.writeheader()
		for inp in inputs:
			logging.info("Scanning %s", inp)
			matches = process_file(inp, compiled, writer, password=args.password, progress_interval=args.progress_interval)
			total_matches += matches
			logging.info("Finished %s: %d matches found", inp, matches)

	logging.info("All done. Total matches written: %d", total_matches)
	return 0


if __name__ == "__main__":
	raise SystemExit(main())

