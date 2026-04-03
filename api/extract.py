from flask import Flask, request, jsonify
import fitz
import re
import json

app = Flask(__name__)

SHEET_RE = re.compile(r'\b([A-Z]{1,3}-\d{1,4}[A-Z]?|[A-Z]{1,3}\d{3,4}[A-Z]?)\b')

NOISE_RE = re.compile(
    r'PROJECT\s*#|PHONE\s*:|FAX\s*:|WWW\.|\.COM|SCALE[\s:=]'
    r'|COPYRIGHT|^\s*\d+\s*$|\d{3}[-.\s]\d{3}[-.\s]\d{4}',
    re.IGNORECASE
)

DISCIPLINE_MAP = {
    'S':  ('STRUCTURAL',          300), 'SD': ('STRUCTURAL',          300),
    'SS': ('STRUCTURAL',          300),
    'C':  ('CIVIL',               200), 'L':  ('LANDSCAPE',           150),
    'I':  ('INTERIORS',           250),
    'P':  ('PLUMBING',            400), 'MP': ('MECHANICAL PLUMBING',  450),
    'M':  ('MECHANICAL',          500), 'H':  ('HVAC',                520),
    'E':  ('ELECTRICAL',          600), 'EP': ('ELECTRICAL POWER',    620),
    'FP': ('FIRE PROTECTION',     700), 'FA': ('FIRE ALARM',          710),
    'SP': ('FIRE SUPPRESSION',    720),
    'T':  ('TECHNOLOGY',          800), 'IT': ('TECHNOLOGY',          800),
}

AUTO_REGIONS = [
    (0.45, 0.55, 1.0, 1.0),
    (0.0,  0.55, 0.55, 1.0),
    (0.0,  0.0,  1.0,  1.0),
]


def resolve_params(num):
    m = re.match(r'^([A-Z]+)', num)
    prefix = num.split('-')[0] if '-' in num else (m.group(1) if m else '')
    key = next(
        (k for k in sorted(DISCIPLINE_MAP, key=len, reverse=True) if prefix == k),
        None
    )
    discipline, major = DISCIPLINE_MAP.get(key, ('CONSULTANT', 900))
    n_match = re.search(r'\d+', num)
    n = int(n_match.group()) if n_match else 0
    minor = major + (10 if n < 100 else 20 if n < 500 else 30)
    return discipline, major, minor


def extract_sheets(pdf_bytes, number_region=None, title_region=None):
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    results = []
    seen = set()

    for page in doc:
        w, h = page.rect.width, page.rect.height
        num = None
        title = ''
        num_block_idx = None
        num_blocks = None

        # ── Find sheet number ──────────────────────────────────────────────
        search_regions = [number_region] if number_region else AUTO_REGIONS
        for (rx0, ry0, rx1, ry1) in search_regions:
            clip = fitz.Rect(rx0 * w, ry0 * h, rx1 * w, ry1 * h)
            blocks = sorted(
                [b for b in page.get_text('blocks', clip=clip) if b[6] == 0],
                key=lambda b: b[1]
            )
            for i, blk in enumerate(blocks):
                m = SHEET_RE.search(blk[4])
                if not m or m.group(0) in seen:
                    continue
                num = m.group(0)
                num_block_idx = i
                num_blocks = blocks
                break
            if num:
                break

        if not num:
            continue

        # ── Extract title ──────────────────────────────────────────────────
        if title_region:
            (tx0, ty0, tx1, ty1) = title_region
            tclip = fitz.Rect(tx0 * w, ty0 * h, tx1 * w, ty1 * h)
            tblocks = sorted(
                [b for b in page.get_text('blocks', clip=tclip) if b[6] == 0],
                key=lambda b: b[1]
            )
            lines = []
            for blk in tblocks:
                for line in blk[4].splitlines():
                    line = line.strip()
                    if line and not NOISE_RE.search(line) and not SHEET_RE.search(line) and len(line) >= 3:
                        lines.append(line)
            title = ' '.join(lines[:3]).strip()
        elif num_blocks and num_block_idx is not None:
            # Heuristic: look in blocks above the sheet number block
            candidates = []
            for j in range(num_block_idx - 1, max(num_block_idx - 12, -1), -1):
                for line in reversed(num_blocks[j][4].splitlines()):
                    line = line.strip()
                    if not line or NOISE_RE.search(line) or SHEET_RE.search(line):
                        continue
                    if len(line) >= 3:
                        candidates.insert(0, line)
                if len(candidates) >= 3:
                    break
            title = ' '.join(candidates[:3]).strip()

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


@app.route('/', methods=['POST'])
@app.route('/api/extract', methods=['POST'])
def extract():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    pdf_bytes = request.files['file'].read()

    number_region = None
    if 'number_region' in request.form:
        try:
            number_region = tuple(json.loads(request.form['number_region']))
        except (ValueError, TypeError):
            pass

    title_region = None
    if 'title_region' in request.form:
        try:
            title_region = tuple(json.loads(request.form['title_region']))
        except (ValueError, TypeError):
            pass

    try:
        sheets = extract_sheets(pdf_bytes, number_region=number_region, title_region=title_region)
        return jsonify({'sheets': sheets})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
