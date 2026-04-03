# -*- coding: utf-8 -*-
"""
extractor.py — Consultant Sheet List Extractor (PyMuPDF)

Extracts sheet numbers and titles from consultant drawing PDFs.
Scans the bottom-right corner of each page (most common title block position),
then bottom-left, then the full page as a fallback.

Usage:
    python extractor.py drawing.pdf               # → CSV to stdout
    python extractor.py drawing.pdf output.csv    # → write to file
"""

import fitz   # pip install pymupdf
import re
import csv
import sys

# ── Patterns ──────────────────────────────────────────────────────────────────
# General: 1-3 uppercase letters + hyphen + 1-4 digits + optional letter
# Examples: E-101, FA-101, FA-1, SD-101, C-001, L-1A
SHEET_RE = re.compile(r'\b([A-Z]{1,3}-\d{1,4}[A-Z]?|[A-Z]{1,3}\d{3,4}[A-Z]?)\b')

NOISE_RE = re.compile(
    r'PROJECT\s*#|PHONE\s*:|FAX\s*:|WWW\.|\.COM|SCALE[\s:=]'
    r'|COPYRIGHT|^\s*\d+\s*$|\d{3}[-.\s]\d{3}[-.\s]\d{4}',
    re.IGNORECASE
)

DISCIPLINE_MAP = {
    'S':  ('STRUCTURAL',         300), 'SD': ('STRUCTURAL',         300),
    'SS': ('STRUCTURAL',         300),
    'C':  ('CIVIL',              200), 'L':  ('LANDSCAPE',          150),
    'I':  ('INTERIORS',          250),
    'P':  ('PLUMBING',           400), 'MP': ('MECHANICAL PLUMBING', 450),
    'M':  ('MECHANICAL',         500), 'H':  ('HVAC',               520),
    'E':  ('ELECTRICAL',         600), 'EP': ('ELECTRICAL POWER',   620),
    'FP': ('FIRE PROTECTION',    700), 'FA': ('FIRE ALARM',         710),
    'SP': ('FIRE SUPPRESSION',   720),
    'T':  ('TECHNOLOGY',         800), 'IT': ('TECHNOLOGY',         800),
}

# Regions to try per page: (x0_frac, y0_frac, x1_frac, y1_frac)
AUTO_REGIONS = [
    (0.45, 0.55, 1.0, 1.0),   # bottom-right (primary)
    (0.0,  0.55, 0.55, 1.0),  # bottom-left  (fallback 1)
    (0.0,  0.0,  1.0,  1.0),  # full page    (fallback 2)
]


def resolve_params(num):
    """Derive discipline, ORDER-MAJOR, and ORDER-MINOR from sheet number prefix."""
    # num is like "FA-101" — split on hyphen for an exact prefix match
    m = re.match(r'^([A-Z]+)', num)
    prefix = num.split('-')[0] if '-' in num else (m.group(1) if m else '')
    key = next(
        (k for k in sorted(DISCIPLINE_MAP, key=len, reverse=True)
         if prefix == k),
        None
    )
    discipline, major = DISCIPLINE_MAP.get(key, ('CONSULTANT', 900))
    n_match = re.search(r'\d+', num)
    n = int(n_match.group()) if n_match else 0
    minor = major + (10 if n < 100 else 20 if n < 500 else 30)
    return discipline, major, minor


def extract_sheets(pdf_path, forced_region=None):
    """
    Extract sheet list from a PDF file.

    Args:
        pdf_path: path to the PDF
        forced_region: optional (x0_frac, y0_frac, x1_frac, y1_frac) to use
                       instead of AUTO_REGIONS (e.g. when user specifies a crop)

    Returns:
        list of dicts with keys: number, title, discipline, order_major, order_minor
    """
    doc     = fitz.open(pdf_path)
    results = []
    seen    = set()

    for page in doc:
        w, h = page.rect.width, page.rect.height
        regions = [forced_region] if forced_region else AUTO_REGIONS

        num   = None
        title = ''

        for (rx0, ry0, rx1, ry1) in regions:
            clip = fitz.Rect(rx0 * w, ry0 * h, rx1 * w, ry1 * h)

            # get_text("blocks") → list of (x0, y0, x1, y1, text, block_no, type)
            # type 0 = text block; sort top-to-bottom by y0
            blocks = sorted(
                [b for b in page.get_text('blocks', clip=clip) if b[6] == 0],
                key=lambda b: b[1]
            )

            # Find the block containing a sheet number
            for i, blk in enumerate(blocks):
                m = SHEET_RE.search(blk[4])
                if not m or m.group(0) in seen:
                    continue

                num = m.group(0)  # full match e.g. "FA-101"

                # Collect candidate title lines from blocks above this one
                candidates = []
                for j in range(i - 1, max(i - 12, -1), -1):
                    for line in reversed(blk[4].splitlines() if j == i else blocks[j][4].splitlines()):
                        line = line.strip()
                        if not line:
                            continue
                        if NOISE_RE.search(line) or SHEET_RE.search(line):
                            continue
                        if len(line) >= 3:
                            candidates.insert(0, line)
                    # Stop if we already have enough
                    if len(candidates) >= 3:
                        break

                title = ' '.join(candidates[:3]).strip()
                break  # found sheet number in this region

            if num:
                break  # found in this region; stop trying others

        if not num:
            continue

        seen.add(num)
        discipline, major, minor = resolve_params(num)
        results.append({
            'number':      num,
            'title':       title,
            'discipline':  discipline,
            'order_major': major,
            'order_minor': minor,
        })

    doc.close()
    return results


def write_csv(sheets, dest):
    """Write sheet list to a CSV file object."""
    writer = csv.writer(dest)
    writer.writerow(['NUMBER', 'SHEET NAME', 'DISCIPLINE', 'ORDER-MAJOR', 'ORDER-MINOR'])
    for s in sheets:
        writer.writerow([s['number'], s['title'], s['discipline'],
                         s['order_major'], s['order_minor']])


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python extractor.py <pdf_path> [output.csv]', file=sys.stderr)
        sys.exit(1)

    pdf_path = sys.argv[1]
    sheets   = extract_sheets(pdf_path)

    if not sheets:
        print('No sheet numbers found.', file=sys.stderr)
        sys.exit(1)

    if len(sys.argv) > 2:
        with open(sys.argv[2], 'w', newline='', encoding='utf-8') as f:
            write_csv(sheets, f)
        print(f'Wrote {len(sheets)} sheets to {sys.argv[2]}', file=sys.stderr)
    else:
        write_csv(sheets, sys.stdout)
