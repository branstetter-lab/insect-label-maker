"""Excel to actual-size insect labels. All PDF measurements are points."""
import argparse
import json
import math
import re
from pathlib import Path
from datetime import datetime, date
from functools import lru_cache

import openpyxl
import zxingcpp
from PIL import Image, ImageOps, ImageFont
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

HEIGHT = .33 * 72
MIN_WIDTH = .68 * 72
SIZE = 3.5
MAIN_SIZE = 4.0
LAYOUT_MODE = "lines"
FONT_PATH = "/System/Library/Fonts/Supplemental/Arial.ttf"
CATALOG_SIZE = 3.75
FIELD_SIZE = 3.1
CATALOG_PAD = .3
PAGE_GAP = .02 * 72
TEXT_SCALE = .75  # condensed text; font height remains unchanged
PAD = .5
BAR = 7.92 * 18 / 14  # original 0.11-inch dark symbol plus two quiet modules per edge
SIDE = 4.43
DEFAULT_LINES = [
    '[Country]: [StateProv], [County]',
    '[LocalityLabel]',
    '[Latitude] [Longitude]±[Uncertainty]m',
    '[Verbatim_Elevation], [Verbatim_Date]',
    '[Collector_Label]',
    '[MacroHabitat_Label]; [Microhabitat_Label]',
]
ALIASES = {'StateProv': 'StateProvince', 'Verbatim_Elevation': 'Verbatium_Elevation'}


def value(v):
    if v is None:
        return ''
    if isinstance(v, (datetime, date)):
        return v.strftime('%d/%b/%Y')
    if isinstance(v, float):
        return format(v, '.12g')
    return str(v).strip().replace('\xa0', ' ')


def read_rows(path, sheet=None):
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = book[sheet] if sheet else book.active
    data = ws.iter_rows(values_only=True)
    headers = [value(x) for x in next(data)]
    if len(set(headers)) != len(headers):
        raise ValueError('Duplicate spreadsheet column names.')
    rows = []
    for number, cells in enumerate(data, 2):
        row = dict(zip(headers, map(value, cells)))
        if not any(row.values()):
            continue
        if not row.get('Catalog_Number'):
            raise ValueError(f'Excel row {number}: Catalog_Number is missing.')
        for alias, source in ALIASES.items():
            row.setdefault(alias, row.get(source, ''))
        if row.get('Collector_Label'):
            row['Collector_Label'] = re.sub(r'\b([A-Z])\. +(?=[A-Z])', r'\1.', row['Collector_Label'])
        row['_excel_row'] = number
        rows.append(row)
    book.close()
    ids = [r['Catalog_Number'] for r in rows]
    if len(ids) != len(set(ids)):
        raise ValueError('Duplicate Catalog_Number values; each specimen must have a unique code.')
    if not rows:
        raise ValueError('No specimen rows found.')
    return rows


def render_lines(row, templates, uncertainty=True):
    import re
    lines = []
    for template in templates:
        if not uncertainty or not row.get('Uncertainty'):
            template = template.replace('±[Uncertainty]m', '')
        keys = re.findall(r'\[([^\]]+)\]', template)
        missing = [k for k in keys if k not in row]
        if missing:
            raise ValueError(f'Unknown template fields: {missing}')
        line = re.sub(r'\[([^\]]+)\]', lambda m: row[m[1]], template)
        lines.append(line.strip(' ,;:'))
    return lines


def width(text, size=SIZE):
    return pdfmetrics.stringWidth(text, 'Arial', size) * TEXT_SCALE


def taxon_italic_ranges(microhabitat):
    """Recognize binomials and Genus sp.; keep qualifiers/families upright."""
    ranges = []
    pattern = r'\b([A-Z][a-z]+)(?:\s+((?:sp|spp)\.|[a-z][a-z-]+))'
    for match in re.finditer(pattern, microhabitat):
        genus, epithet = match.groups()
        if genus.endswith(('aceae', 'idae', 'oideae', 'inae')):
            continue
        if genus.lower() in {'pan', 'flower', 'yellow', 'white', 'purple', 'alpine', 'rocky', 'meadow', 'air', 'ground'}:
            continue
        ranges.append((match.start(1), match.end(1)))
        if epithet not in {'sp.', 'spp.'}:
            ranges.append((match.start(2), match.end(2)))
    return ranges


def label_italic_ranges(lines, microhabitat):
    """Map microhabitat-only emphasis onto possibly wrapped physical lines."""
    joined = ' '.join(lines)
    start = joined.rfind(microhabitat) if microhabitat else -1
    if start < 0:
        return [[] for _ in lines]
    spans = [(start + a, start + b) for a, b in taxon_italic_ranges(microhabitat)]
    result, offset = [], 0
    for line in lines:
        result.append([(max(a, offset) - offset, min(b, offset + len(line)) - offset)
                       for a, b in spans if a < offset + len(line) and b > offset])
        offset += len(line) + 1
    return result


def coordinate_runs(row, line, size):
    coordinates = f"{row.get('Latitude', '')} {row.get('Longitude', '')}"
    uncertainty = f"±{row.get('Uncertainty', '')}m"
    if (LAYOUT_MODE == 'lines' and row.get('Uncertainty')
            and line == coordinates + uncertainty):
        return [(coordinates, size, .62), (uncertainty, 3.5, TEXT_SCALE)]
    return [(line, size, TEXT_SCALE)]


def run_width(runs):
    return sum(pdfmetrics.stringWidth(text, 'Arial', size) * scale for text, size, scale in runs)


def draw_text(pdf, x, y, text, right=False, size=SIZE, italic_ranges=(), horizontal_scale=TEXT_SCALE):
    if right:
        x -= width(text, size)
    obj = pdf.beginText(x, y)
    obj.setFont('Arial', size)
    obj.setHorizScale(horizontal_scale * 100)
    offset = 0
    for start, end in italic_ranges:
        obj.setFont('Arial', size)
        obj.textOut(text[offset:start])
        obj.setFont('ArialItalic', size)
        obj.textOut(text[start:end])
        offset = end
    obj.setFont('Arial', size)
    obj.textOut(text[offset:])
    pdf.drawText(obj)


def wrap(text, available):
    if not text:
        return ['']
    result, current = [], ''
    for word in text.split():
        candidate = f'{current} {word}'.strip()
        if width(candidate) <= available:
            current = candidate
        else:
            if current:
                result.append(current)
            if width(word) > available:
                return None  # widen rather than split specimen data arbitrarily
            current = word
    return result + ([current] if current else [])


LEADING = 3.15
NORMAL_LEADING = 3.8
BAR_Y = (HEIGHT - BAR) / 2
EVENT_Y = .95


def text_size(index):
    return MAIN_SIZE if index < 5 else SIZE


def text_baseline(index):
    if MAIN_SIZE == 4:
        return HEIGHT - .3 - pdfmetrics.getAscent('Arial') * MAIN_SIZE / 1000 - index * NORMAL_LEADING
    return HEIGHT - PAD - SIZE - index * LEADING


@lru_cache(maxsize=8)
def ink_font(size):
    return ImageFont.truetype(FONT_PATH, round(size * 1000))


def ink_bounds(text, size):
    box = ink_font(size).getbbox(text or ' ', anchor='ls')
    return -box[1] / 1000, box[3] / 1000


def line_available(index, label_width, text=None, field_number='MGB0000'):
    return label_width - 2 * PAD - SIDE - BAR - 1.7


class LabelLines(list):
    def __init__(self, lines, sizes, baselines):
        super().__init__(lines)
        self.sizes = sizes
        self.baselines = baselines


def available_at(label_width, baseline, text, size, field_number):
    above, below = ink_bounds(text, size)
    low, high = baseline - below, baseline + above
    available = label_width - 2 * PAD - SIDE
    # Shared invisible text boundary, including lines above/below the code.
    return available - BAR - 1.7


def flow_layout(row, templates, label_width, uncertainty):
    rendered = render_lines(row, templates, uncertainty)
    # Tag habitat words so any physical line containing them stays at 3.5 pt.
    words = []
    for i, field in enumerate(rendered):
        if not field:
            continue
        parts = field.replace(';', ',').split()
        if words:
            old, tag = words[-1]
            words[-1] = (old.rstrip(',') + ',', tag)
        words.extend((word, i >= 5) for word in parts)

    @lru_cache(maxsize=None)
    def fit(start, top_milli, depth):
        if start == len(words):
            return (), (), ()
        if depth >= 7:
            return None
        top = top_milli / 1000
        for size in ([4, 3.5] if MAIN_SIZE == 4 else [3.5]):
            candidates = []
            for end in range(start + 1, len(words) + 1):
                if size == 4 and words[end - 1][1]:
                    break
                text = ' '.join(word for word, _ in words[start:end])
                above, below = ink_bounds(text, size)
                baseline = top - above
                if baseline - below < .2:
                    break
                available = available_at(label_width, baseline, text, size, row.get('Field_Number', ''))
                if width(text, size) <= available:
                    candidates.append((end, text, baseline, below))
            for end, text, baseline, below in reversed(candidates):
                rest = fit(end, round((baseline - below - .22) * 1000), depth + 1)
                if rest is not None:
                    return ((text,) + rest[0], (size,) + rest[1], (baseline,) + rest[2])
        return None

    result = fit(0, round((HEIGHT - .3) * 1000), 0)
    return LabelLines(*result) if result is not None else None


def _layout(row, templates, label_width, uncertainty):
    if width(row.get('Field_Number', ''), FIELD_SIZE) > BAR + 7:
        return None
    if width(row['Catalog_Number'], CATALOG_SIZE) > HEIGHT - 2 * CATALOG_PAD:
        return None
    if LAYOUT_MODE == 'flow':
        return flow_layout(row, templates, label_width, uncertainty)
    if MAIN_SIZE == 3.5:
        # Preserve the previously printed 3.5-point profile.
        result = []
        for line in render_lines(row, templates, uncertainty):
            words = line.split()
            if not words:
                result.append('')
            while words:
                available = line_available(len(result), label_width, line, row.get('Field_Number', ''))
                current = ''
                while words and width((current + ' ' + words[0]).strip()) <= available:
                    current = (current + ' ' + words.pop(0)).strip()
                if not current:
                    return None
                result.append(current)
        if text_baseline(len(result) - 1) - .8 < PAD:
            return None
        return LabelLines(result, [3.5] * len(result), [text_baseline(i) for i in range(len(result))])
    result, sizes, baselines = [], [], []
    previous_bottom = HEIGHT - .3
    for i, line in enumerate(render_lines(row, templates, uncertainty)):
        chosen = None
        for size in ([4, 3.5] if i < 5 else [3.5]):
            above, below = ink_bounds(line, size)
            baseline = text_baseline(i)
            available = available_at(label_width, baseline, line, size, row.get('Field_Number', ''))
            if (run_width(coordinate_runs(row, line, size)) <= available and baseline + above <= previous_bottom + .001
                    and baseline - below >= .1):
                chosen = size
                break
        if chosen is None:
            return None
        result.append(line)
        sizes.append(chosen)
        baselines.append(baseline)
        previous_bottom = baseline - ink_bounds(line, chosen)[1]
    return LabelLines(result, sizes, baselines)


HABITAT_TEMPLATE = '[MacroHabitat_Label]; [Microhabitat_Label]'


def layout(row, templates, label_width, uncertainty):
    full = _layout(row, templates, label_width, uncertainty)
    if full is not None:
        return full
    # Prefer removing uncertainty before removing macrohabitat.
    if uncertainty and _layout(row, templates, label_width, False) is not None:
        return None
    if (HABITAT_TEMPLATE not in templates or not row.get('Microhabitat_Label')
            or not row.get('MacroHabitat_Label')):
        return None
    reduced = [t.replace(HABITAT_TEMPLATE, '[Microhabitat_Label]') for t in templates]
    return _layout(row, reduced, label_width, uncertainty)


@lru_cache(maxsize=4096)
def matrix(code):
    barcode = zxingcpp.create_barcode(code, zxingcpp.BarcodeFormat.DataMatrix)
    img = Image.fromarray(zxingcpp.write_barcode_to_image(barcode, scale=1, add_quiet_zones=True))
    decoded = zxingcpp.read_barcode(img)
    if decoded is None or decoded.text != code:
        raise ValueError(f'Barcode verification failed for {code}')
    return ImageOps.expand(img.convert('L'), border=1, fill=255)


def draw_label(pdf, row, lines, x, y, label_width, guides=False):
    if guides:
        pdf.setStrokeColorRGB(.8, .8, .8)
        pdf.setLineWidth(.15)
        pdf.rect(x, y, label_width, HEIGHT)
        pdf.setStrokeColorRGB(0, 0, 0)
    pdf.setFont('Arial', SIZE)
    italics = label_italic_ranges(lines, row.get('Microhabitat_Label', ''))
    for i, line in enumerate(lines):
        runs = coordinate_runs(row, line, lines.sizes[i])
        if len(runs) == 1:
            draw_text(pdf, x + PAD, y + lines.baselines[i], line, size=lines.sizes[i], italic_ranges=italics[i])
        else:
            cursor = x + PAD
            for text, size, scale in runs:
                draw_text(pdf, cursor, y + lines.baselines[i], text, size=size, horizontal_scale=scale)
                cursor += pdfmetrics.stringWidth(text, 'Arial', size) * scale
    pdf.saveState()
    pdf.translate(x + label_width - PAD - .75, y + (HEIGHT - width(row['Catalog_Number'], CATALOG_SIZE)) / 2)
    pdf.rotate(90)
    draw_text(pdf, 0, 0, row['Catalog_Number'], size=CATALOG_SIZE)
    pdf.restoreState()
    bx = x + label_width - PAD - SIDE - BAR
    by = y + BAR_Y
    img = matrix(row['Catalog_Number'])
    # Vector modules preserve sharp edges at printer resolution.
    step = BAR / max(img.size)
    for iy in range(img.height):
        for ix in range(img.width):
            if img.getpixel((ix, iy)) < 128:
                pdf.rect(bx + ix * step, by + (img.height - 1 - iy) * step,
                         step, step, fill=1, stroke=0)
    draw_text(pdf, bx + BAR, y + EVENT_Y, row.get('Field_Number', ''), right=True, size=FIELD_SIZE)


def generate(args):
    global MAIN_SIZE, FONT_PATH, LAYOUT_MODE
    MAIN_SIZE = args.main_size
    LAYOUT_MODE = args.layout
    FONT_PATH = args.font
    ink_font.cache_clear()
    font = Path(args.font)
    if not font.is_file():
        raise ValueError('Arial font not found. Supply its .ttf path with --font.')
    pdfmetrics.registerFont(TTFont('Arial', str(font)))
    italic_path = Path(args.italic_font)
    if not italic_path.is_file():
        raise ValueError('Arial Italic font not found. Supply --italic-font with its .ttf path.')
    pdfmetrics.registerFont(TTFont('ArialItalic', str(italic_path)))
    # Arial regular and italic share advance widths; preserve fit measurements.
    for code in range(32, 127):
        if abs(pdfmetrics.stringWidth(chr(code), 'Arial', SIZE) - pdfmetrics.stringWidth(chr(code), 'ArialItalic', SIZE)) > .001:
            raise ValueError('Regular and italic font widths differ; use matching Arial font files.')
    rows = read_rows(args.excel, args.sheet)
    templates = json.loads(Path(args.template).read_text()) if args.template else DEFAULT_LINES
    if not isinstance(templates, list) or not templates or not all(isinstance(t, str) for t in templates):
        raise ValueError('Template must be a JSON list of line strings.')
    selected = []
    for row in rows:
        found = None
        # Prefer the requested width, first with uncertainty, then without it.
        for w in range(math.ceil(args.width * 7200), math.floor(args.max_width * 7200) + 1, 72):
            label_width = w / 100
            for uncertainty in (True, False):
                lines = layout(row, templates, label_width, uncertainty)
                if lines is not None:
                    found = (label_width, uncertainty, lines)
                    break
            if found:
                break
        if found is None:
            raise ValueError(f"Excel row {row['_excel_row']} ({row['Catalog_Number']}) cannot fit at {MAIN_SIZE:g} pt main / 3.5 pt habitat within {args.max_width} inches. Edit its label fields or raise --max-width.")
        selected.append(found)
    common_width = max(t[0] for t in selected)
    # Use one consistent cutting grid and restore uncertainty wherever it now fits.
    plans = []
    for row in rows:
        lines = layout(row, templates, common_width, True)
        keep = lines is not None
        if lines is None:
            lines = layout(row, templates, common_width, False)
        plans.append((lines, keep))
    margin, gap = 18, PAGE_GAP
    cols = int((612 - 2 * margin + gap) // (common_width + gap))
    per_col = int((792 - 2 * margin + gap) // (HEIGHT + gap))
    if cols < 1:
        raise ValueError('Label width exceeds printable page width.')
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    temp = output.with_suffix('.tmp.pdf')
    pdf = canvas.Canvas(str(temp), pagesize=(612, 792))
    pdf.setTitle('Insect specimen labels')
    report = []
    for index, (row, (lines, keep)) in enumerate(zip(rows, plans)):
        position = index % (cols * per_col)
        if index and position == 0:
            pdf.showPage()
        col, r = divmod(position, per_col)
        x = margin + col * (common_width + gap)
        y = 792 - margin - HEIGHT - r * (HEIGHT + gap)
        draw_label(pdf, row, lines, x, y, common_width, args.guides)
        report.append({'excel_row': row['_excel_row'], 'catalog_number': row['Catalog_Number'],
                       'uncertainty_omitted': bool(row.get('Uncertainty')) and not keep,
                       'macrohabitat_omitted': _layout(row, templates, common_width, keep) is None,
                       'line_text_runs': [coordinate_runs(row, text, size) for text, size in zip(lines, lines.sizes)],
                       'line_font_points': lines.sizes, 'line_baselines': lines.baselines, 'lines': lines})
    pdf.save()
    temp.replace(output)
    summary = {'labels': len(rows), 'width_inches': common_width / 72, 'height_inches': .33,
               'layout': LAYOUT_MODE, 'font': 'Arial', 'font_points': MAIN_SIZE, 'habitat_font_points': SIZE, 'catalog_font_points': CATALOG_SIZE, 'field_font_points': FIELD_SIZE,
               'barcode_dark_symbol_inches': 0.11, 'barcode_center_y_inches': HEIGHT / 144,
               'horizontal_gap_inches': gap / 72, 'vertical_gap_inches': gap / 72, 'text_horizontal_scale': TEXT_SCALE, 'barcode': 'Data Matrix',
               'macrohabitat_omitted_count': sum(r['macrohabitat_omitted'] for r in report),
               'uncertainty_omitted_count': sum(r['uncertainty_omitted'] for r in report), 'records': report}
    output.with_suffix('.json').write_text(json.dumps(summary, indent=2))
    print(f"Created {output}: {len(rows)} labels, {common_width / 72:.2f} × 0.33 in, Arial {MAIN_SIZE:g} pt main / 3.5 pt habitat.")
    print(f"Macrohabitat omitted on {summary['macrohabitat_omitted_count']} labels.")
    print(f"Uncertainty omitted on {summary['uncertainty_omitted_count']} labels. Review {output.with_suffix('.json')}.")


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('excel')
    p.add_argument('-o', '--output', default='output/labels.pdf')
    p.add_argument('--sheet')
    p.add_argument('--layout', choices=['lines', 'flow'], default='lines', help='Separate fields into lines, or wrap comma-separated fields')
    p.add_argument('--italic-font', default='/System/Library/Fonts/Supplemental/Arial Italic.ttf')
    p.add_argument('--main-size', type=float, choices=[3.5, 4.0], default=4.0, help='Font size for the first five physical text lines')
    p.add_argument('--template', help='JSON list of lines with [Excel_Column] placeholders')
    p.add_argument('--font', default='/System/Library/Fonts/Supplemental/Arial.ttf')
    p.add_argument('--width', type=float, default=.68, help='Starting width in inches')
    p.add_argument('--max-width', type=float, default=.68, help='Maximum width in inches (raise explicitly to allow widening)')
    p.add_argument('--guides', action='store_true', help='Add light cutting rectangles')
    args = p.parse_args()
    if not MIN_WIDTH / 72 <= args.width <= args.max_width <= 8:
        p.error('Widths must satisfy 0.68 <= width <= max-width <= 8 inches.')
    try:
        generate(args)
    except (ValueError, KeyError, OSError) as exc:
        p.exit(1, f'Error: {exc}\n')

if __name__ == '__main__':
    main()
