"""Excel to actual-size insect labels. All PDF measurements are points."""
import argparse
import json
import math
import re
import io
import shutil
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime, date
from functools import lru_cache

import openpyxl
import zxingcpp
from PIL import Image, ImageOps, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

HEIGHT = .33 * 72
MAX_WIDTH_INCHES = .68
MIN_WIDTH = MAX_WIDTH_INCHES * 72  # Default width retained for API compatibility.
SIZE = 3.5
MAIN_SIZE = 4.0
LAYOUT_MODE = "lines"
FONT_PATH = "/System/Library/Fonts/Supplemental/Arial.ttf"
ITALIC_FONT_PATH = "/System/Library/Fonts/Supplemental/Arial Italic.ttf"
CATALOG_SIZE = 3.75
FIELD_SIZE = 4.0
FIELD_FONT = "Arial"
FIELD_SCALE = .60
FIELD_LEFT_SHIFT = 0.0
HABITAT_FIELD_GAP = 1.5
HEADING_HEIGHT = 8.0
CATALOG_PAD = .3
PAGE_GAP = .02 * 72
TEXT_SCALE = .75  # condensed text; font height remains unchanged
COORDINATE_SCALE = .61
PAD = .5
BAR = 7.92 * 18 / 14 * 1.02  # original 0.11-inch dark symbol plus two quiet modules per edge
BARCODE_RENDERING = 'raster'
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
    return pdfmetrics.stringWidth(text, 'LabelRegular', size) * TEXT_SCALE


def fit_width(text, size=SIZE):
    return sum(max(pdfmetrics.stringWidth(char, 'LabelRegular', size),
                   pdfmetrics.stringWidth(char, 'LabelItalic', size)) for char in text) * TEXT_SCALE


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
        return [(coordinates, size, COORDINATE_SCALE), (uncertainty, 3.5, TEXT_SCALE)]
    return [(line, size, TEXT_SCALE)]


def field_width(text):
    return pdfmetrics.stringWidth(text, FIELD_FONT, FIELD_SIZE) * FIELD_SCALE


def run_width(runs):
    return sum(pdfmetrics.stringWidth(text, 'LabelRegular', size) * scale for text, size, scale in runs)


def draw_text(pdf, x, y, text, right=False, size=SIZE, italic_ranges=(), horizontal_scale=None):
    if right:
        x -= width(text, size)
    obj = pdf.beginText(x, y)
    obj.setFont('LabelRegular', size)
    obj.setCharSpace(0)
    obj.setHorizScale((TEXT_SCALE if horizontal_scale is None else horizontal_scale) * 100)
    offset = 0
    for start, end in italic_ranges:
        obj.setFont('LabelRegular', size)
        obj.textOut(text[offset:start])
        obj.setFont('LabelItalic', size)
        obj.textOut(text[start:end])
        offset = end
    obj.setFont('LabelRegular', size)
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
        return HEIGHT - .3 - pdfmetrics.getAscent('LabelRegular') * MAIN_SIZE / 1000 - index * NORMAL_LEADING
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


def habitat_available(label_width, field_number):
    boundary = label_width - 2 * PAD - SIDE - BAR - 1.7
    if not field_number:
        return boundary
    field_left = label_width - PAD - SIDE - field_width(field_number) - FIELD_LEFT_SHIFT
    return min(boundary, field_left - HABITAT_FIELD_GAP - PAD)


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
        if depth >= max(7, int(HEIGHT / 3)):
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
                if any(is_habitat for _, is_habitat in words[start:end]):
                    available = min(available, habitat_available(label_width, row.get('Field_Number', '')))
                if fit_width(text, size) <= available:
                    candidates.append((end, text, baseline, below))
            for end, text, baseline, below in reversed(candidates):
                rest = fit(end, round((baseline - below - .22) * 1000), depth + 1)
                if rest is not None:
                    return ((text,) + rest[0], (size,) + rest[1], (baseline,) + rest[2])
        return None

    result = fit(0, round((HEIGHT - .3) * 1000), 0)
    return LabelLines(*result) if result is not None else None


def _layout(row, templates, label_width, uncertainty):
    if field_width(row.get('Field_Number', '')) + FIELD_LEFT_SHIFT > BAR + 1.7:
        return None
    if width(row['Catalog_Number'], CATALOG_SIZE) > HEIGHT - 2 * CATALOG_PAD:
        return None
    if LAYOUT_MODE == 'flow':
        return flow_layout(row, templates, label_width, uncertainty)
    if MAIN_SIZE == 3.5:
        # Preserve the previously printed 3.5-point profile.
        result = []
        for source_index, line in enumerate(render_lines(row, templates, uncertainty)):
            words = line.split()
            if not words:
                result.append('')
            while words:
                available = line_available(len(result), label_width, line, row.get('Field_Number', ''))
                if source_index >= 5:
                    available = min(available, habitat_available(label_width, row.get('Field_Number', '')))
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
            if i >= 5:
                available = min(available, habitat_available(label_width, row.get('Field_Number', '')))
            if ((fit_width(line, size) if i >= 5 else run_width(coordinate_runs(row, line, size))) <= available and baseline + above <= previous_bottom + .001
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
    pdf.setFont('LabelRegular', SIZE)
    italics = label_italic_ranges(lines, row.get('Microhabitat_Label', ''))
    for i, line in enumerate(lines):
        runs = coordinate_runs(row, line, lines.sizes[i])
        if len(runs) == 1:
            draw_text(pdf, x + PAD, y + lines.baselines[i], line, size=lines.sizes[i], italic_ranges=italics[i])
        else:
            cursor = x + PAD
            for text, size, scale in runs:
                draw_text(pdf, cursor, y + lines.baselines[i], text, size=size, horizontal_scale=scale)
                cursor += pdfmetrics.stringWidth(text, 'LabelRegular', size) * scale
    pdf.saveState()
    pdf.translate(x + label_width - PAD - .75, y + (HEIGHT - width(row['Catalog_Number'], CATALOG_SIZE)) / 2)
    pdf.rotate(90)
    draw_text(pdf, 0, 0, row['Catalog_Number'], size=CATALOG_SIZE)
    pdf.restoreState()
    bx = x + label_width - PAD - SIDE - BAR
    by = y + BAR_Y
    img = matrix(row['Catalog_Number'])
    if BARCODE_RENDERING == 'raster':
        pixels = round(BAR / 72 * 1200)
        raster = img.resize((pixels, pixels), Image.Resampling.NEAREST)
        pdf.drawImage(ImageReader(raster), bx, by, width=BAR, height=BAR)
    else:
        draw_vector_matrix(pdf, img, bx, by)
    event = row.get('Field_Number', '')
    obj = pdf.beginText(bx + BAR - field_width(event) - FIELD_LEFT_SHIFT, y + EVENT_Y)
    obj.setFont(FIELD_FONT, FIELD_SIZE)
    obj.setCharSpace(0)
    obj.setHorizScale(FIELD_SCALE * 100)
    obj.textOut(event)
    pdf.drawText(obj)


def draw_vector_matrix(pdf, img, bx, by):
    # Vector modules preserve sharp edges at printer resolution.
    step = BAR / max(img.size)
    for iy in range(img.height):
        for ix in range(img.width):
            if img.getpixel((ix, iy)) < 128:
                pdf.rect(bx + ix * step, by + (img.height - 1 - iy) * step,
                         step, step, fill=1, stroke=0)


def separate_habitat_templates(templates):
    habitat_fields = {'MacroHabitat_Label', 'Microhabitat_Label'}
    result = []
    for template in templates:
        fields = set(re.findall(r'\[([^\]]+)\]', template))
        if fields & habitat_fields:
            if fields - habitat_fields:
                raise ValueError('Separate habitat labels require habitat fields in their own template entries.')
            continue
        result.append(template)
    if not result:
        raise ValueError('The main label template has no fields left after separating habitat.')
    return result


def habitat_label_plan(row, label_width, height, max_lines=3):
    fields = [' '.join(row[key].split()) for key in ('MacroHabitat_Label', 'Microhabitat_Label')]
    texts, sources = [], []
    available = label_width - 2 * PAD - .3
    for field, source in zip(fields, ['macrohabitat', 'microhabitat']):
        current = ''
        for word in field.split():
            candidate = (current + ' ' + word).strip()
            if fit_width(candidate, 4) <= available:
                current = candidate
            else:
                if current:
                    texts.append(current)
                    sources.append(source)
                if fit_width(word, 4) > available:
                    raise ValueError(f"Excel row {row['_excel_row']} ({row['Catalog_Number']}): habitat word cannot fit at 4 pt: {word!r}. Shorten its label field.")
                current = word
        texts.append(current)
        sources.append(source)
    if len(texts) > max_lines:
        raise ValueError(f"Excel row {row['_excel_row']} ({row['Catalog_Number']}): habitat needs {len(texts)} lines; maximum is {max_lines}. Shorten its label fields.")
    emphasis = label_italic_ranges(texts, fields[1])
    bounds = []
    italic = ImageFont.truetype(ITALIC_FONT_PATH, 4000)
    for text, spans in zip(texts, emphasis):
        above, below = ink_bounds(text or 'Ag', 4)
        if spans:
            box = italic.getbbox(text, anchor='ls')
            above, below = max(above, -box[1] / 1000), max(below, box[3] / 1000)
        bounds.append((above, below))
    leading = max([4.6] + [bounds[i][1] + bounds[i + 1][0] + .3 for i in range(len(texts) - 1)])
    block_height = bounds[0][0] + leading * (len(texts) - 1) + bounds[-1][1]
    if block_height > height - 2 * PAD:
        raise ValueError('Habitat label is too short for its 4-point lines. Increase --habitat-height.')
    top_baseline = (height + block_height) / 2 - bounds[0][0]
    return {'lines': texts, 'line_sources': sources, 'italic_ranges': emphasis, 'font_points': 4,
            'baselines': [top_baseline - i * leading for i in range(len(texts))],
            'height_inches': height / 72, 'width_inches': label_width / 72}


def draw_habitat_label(pdf, plan, x, y, label_width, guides):
    if guides:
        pdf.setStrokeColorRGB(.8, .8, .8)
        pdf.setLineWidth(.15)
        pdf.rect(x, y, label_width, plan['height_inches'] * 72)
        pdf.setStrokeColorRGB(0, 0, 0)
    for i, text in enumerate(plan['lines']):
        emphasis = plan['italic_ranges'][i]
        draw_text(pdf, x + PAD, y + plan['baselines'][i], text, size=4, italic_ranges=emphasis)


def configure_dimensions(width_inches, height_inches, max_width=None):
    global HEIGHT, BAR_Y
    if not math.isfinite(width_inches) or not 0 < width_inches <= MAX_WIDTH_INCHES:
        raise ValueError('Label width must be greater than 0 and never exceed 0.68 inches (four-row unit-tray limit).')
    if max_width is not None and (not math.isfinite(max_width) or max_width != width_inches):
        raise ValueError('--max-width no longer allows widening; omit it or set it equal to --width (at most 0.68 inches).')
    if not math.isfinite(height_inches) or not .33 <= height_inches <= 10.5:
        raise ValueError('Label height must be between 0.33 and 10.5 inches, within a Letter page.')
    HEIGHT = height_inches * 72
    BAR_Y = (HEIGHT - BAR) / 2


def sheet_positions(rows, label_width, pair_height, headings=True):
    """Pack intact pairs down columns; repeat field headings at continuations."""
    margin, gap = 18, PAGE_GAP
    columns = int((612 - 2 * margin + gap) // (label_width + gap))
    if columns < 1:
        raise ValueError('Label width exceeds printable page width.')
    page, column, top = 1, 0, 792 - margin
    previous_field = None
    result = []
    for row in rows:
        field = row.get('Field_Number', '')
        while True:
            needs_heading = headings and bool(field) and (field != previous_field or top == 792 - margin)
            reserve = HEADING_HEIGHT if needs_heading else 0
            if top - reserve - pair_height >= margin - .0001:
                break
            if top == 792 - margin:
                raise ValueError('The label pair and its field heading cannot fit on a page. Reduce label heights or gap.')
            column += 1
            if column == columns:
                column, page = 0, page + 1
            top = 792 - margin
        x = margin + column * (label_width + gap)
        heading = {'text': f'---- {field} ----', 'field_number': field, 'continued': field == previous_field,
                   'x_points': x, 'baseline': top - 5} if needs_heading else None
        top -= reserve
        result.append({'page': page, 'x_points': x, 'y_points': top - HEIGHT, 'heading': heading})
        top -= pair_height + gap
        previous_field = field
    return result


def heading_segments(field_number, label_width):
    size, dash_size, gap = 5, 3.5, 1
    available = label_width - 2 * PAD
    number_width = pdfmetrics.stringWidth(field_number, 'FieldHeading', size)
    scale = min(1, (available - 2 * gap - 4) / number_width)
    number_width *= scale
    side_width = (available - number_width - 2 * gap) / 2
    dash_width = pdfmetrics.stringWidth('-', 'FieldHeading', dash_size)
    count = max(2, int(side_width / dash_width))
    return {'size': size, 'dash_size': dash_size, 'number_scale': scale,
            'number_width': number_width, 'side_width': side_width,
            'dash_count': count, 'dash_spacing': (side_width - count * dash_width) / (count - 1), 'gap': gap}


def draw_field_heading(pdf, heading, label_width):
    pdf.saveState()
    field = heading['field_number']
    parts = heading_segments(field, label_width)
    x = heading['x_points'] + PAD
    number_x = x + parts['side_width'] + parts['gap']
    obj = pdf.beginText(number_x, heading['baseline'])
    obj.setFont('FieldHeading', parts['size'])
    obj.setHorizScale(parts['number_scale'] * 100)
    obj.textOut(field)
    pdf.drawText(obj)
    for dash_x in [x, number_x + parts['number_width'] + parts['gap']]:
        obj = pdf.beginText(dash_x, heading['baseline'] + .6)
        obj.setFont('FieldHeading', parts['dash_size'])
        obj.setCharSpace(parts['dash_spacing'])
        obj.textOut('-' * parts['dash_count'])
        pdf.drawText(obj)
    heading['text'] = '-' * parts['dash_count'] + ' ' + field + ' ' + '-' * parts['dash_count']
    pdf.restoreState()



def draw_raster_label(pdf, draw, x, y, label_width, label_height):
    """Embed one complete label image at 1200 dpi, without a text layer."""
    with tempfile.TemporaryDirectory(prefix='label-raster-') as directory:
        source = Path(directory) / 'label.pdf'
        small = canvas.Canvas(str(source), pagesize=(label_width, label_height))
        draw(small)
        small.save()
        result = subprocess.run(
            ['pdftoppm', '-r', '1200', '-singlefile', '-png', str(source)],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
        img = Image.open(io.BytesIO(result.stdout)).convert('RGB')
        pdf.drawImage(ImageReader(img), x, y, width=label_width, height=label_height)


def generate(args):
    global BARCODE_RENDERING, BAR
    barcode_scale = getattr(args, 'barcode_scale', 1.02)
    if not math.isfinite(barcode_scale) or not 1 <= barcode_scale <= 1.1:
        raise ValueError('--barcode-scale must be between 1.0 and 1.1.')
    BAR = 7.92 * 18 / 14 * barcode_scale
    BARCODE_RENDERING = getattr(args, 'barcode_rendering', 'raster')
    label_rendering = getattr(args, 'label_rendering', 'vector')
    if label_rendering == 'raster' and not shutil.which('pdftoppm'):
        raise ValueError('Whole-label raster rendering requires pdftoppm (Poppler). On macOS: brew install poppler.')
    global MAIN_SIZE, FONT_PATH, LAYOUT_MODE, ITALIC_FONT_PATH, FIELD_SIZE, FIELD_FONT, TEXT_SCALE, COORDINATE_SCALE
    configure_dimensions(args.width, args.height, args.max_width)
    FIELD_SIZE = args.field_size
    MAIN_SIZE = args.main_size
    LAYOUT_MODE = args.layout
    family = 'Verdana' if args.font_family == 'verdana' else 'Arial'
    TEXT_SCALE = .70 if family == 'Verdana' else .75
    COORDINATE_SCALE = .49 if family == 'Verdana' else .61
    coordinate_scale = getattr(args, 'coordinate_scale', None)
    if coordinate_scale is not None:
        if not math.isfinite(coordinate_scale) or not 0 < coordinate_scale <= 1:
            raise ValueError('--coordinate-scale must be greater than zero and at most 1.')
        COORDINATE_SCALE = coordinate_scale
    font = Path(args.font) if args.font else Path('/System/Library/Fonts/Supplemental') / (family + '.ttf')
    FONT_PATH = str(font)
    ink_font.cache_clear()
    if not font.is_file():
        raise ValueError(f'{family} regular font not found. Supply its .ttf path with --font.')
    pdfmetrics.registerFont(TTFont('LabelRegular', str(font)))
    italic_path = Path(args.italic_font) if args.italic_font else font.with_name(family + ' Italic.ttf')
    ITALIC_FONT_PATH = str(italic_path)
    if not italic_path.is_file():
        raise ValueError('Italic font not found. Supply a matching .ttf path with --italic-font.')
    pdfmetrics.registerFont(TTFont('LabelItalic', str(italic_path)))
    field_path = Path(args.field_font) if args.field_font else font.with_name('Verdana.ttf')
    if not field_path.is_file():
        if args.field_font:
            raise ValueError('Field-number font not found. Supply --field-font with a valid .ttf path.')
        field_path = font
    pdfmetrics.registerFont(TTFont('FieldNumber', str(field_path)))
    FIELD_FONT = 'FieldNumber'
    heading_path = Path(args.heading_font) if args.heading_font else font.with_name(family + ' Bold.ttf')
    if not heading_path.is_file():
        heading_path = font.with_name('verdanab.ttf' if family == 'Verdana' else 'arialbd.ttf')
    if not heading_path.is_file() and not args.no_field_headings:
        raise ValueError('Bold heading font not found. Supply --heading-font with a matching bold .ttf path.')
    if heading_path.is_file():
        pdfmetrics.registerFont(TTFont('FieldHeading', str(heading_path)))
    rows = read_rows(args.excel, args.sheet)
    templates = json.loads(Path(args.template).read_text()) if args.template else DEFAULT_LINES
    if not isinstance(templates, list) or not templates or not all(isinstance(t, str) for t in templates):
        raise ValueError('Template must be a JSON list of line strings.')
    if not math.isfinite(args.habitat_height) or args.habitat_height <= 0:
        raise ValueError('--habitat-height must be a positive number of inches.')
    if not math.isfinite(args.habitat_gap) or args.habitat_gap < 0:
        raise ValueError('--habitat-gap must be zero or a positive number of inches.')
    if args.separate_habitat:
        if not math.isfinite(args.habitat_width) or not 0 < args.habitat_width < args.width:
            raise ValueError('--habitat-width must be positive and smaller than the main-label width.')
        for row in rows:
            if not {'MacroHabitat_Label', 'Microhabitat_Label'} <= row.keys():
                raise ValueError('Separate habitat labels require MacroHabitat_Label and Microhabitat_Label columns.')
        templates = separate_habitat_templates(templates)
    common_width = args.width * 72
    habitat_height = args.habitat_height * 72
    habitat_gap = args.habitat_gap * 72
    habitat_plans = [habitat_label_plan(row, args.habitat_width * 72, habitat_height, args.habitat_max_lines) for row in rows] if args.separate_habitat else []
    for row in rows:
        if not any(layout(row, templates, common_width, uncertainty) is not None for uncertainty in (True, False)):
            raise ValueError(f"Excel row {row['_excel_row']} ({row['Catalog_Number']}) cannot fit at {MAIN_SIZE:g} pt main / 3.5 pt habitat within {args.width:g} × {args.height:g} inches. Shorten label fields, try --layout flow, or increase --height. Width cannot exceed 0.68 inches.")
    # Use one consistent cutting grid and restore uncertainty wherever it now fits.
    plans = []
    for row in rows:
        lines = layout(row, templates, common_width, True)
        keep = lines is not None
        if lines is None:
            lines = layout(row, templates, common_width, False)
        plans.append((lines, keep))
    margin, gap = 18, PAGE_GAP
    pair_height = HEIGHT + habitat_gap + habitat_height if args.separate_habitat else HEIGHT
    positions = sheet_positions(rows, common_width, pair_height, not args.no_field_headings)
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    temp = output.with_suffix('.tmp.pdf')
    pdf = canvas.Canvas(str(temp), pagesize=(612, 792))
    pdf.setTitle('Insect specimen labels')
    report = []
    current_page = 1
    for index, (row, (lines, keep)) in enumerate(zip(rows, plans)):
        placement = positions[index]
        if placement['page'] != current_page:
            pdf.showPage()
            current_page = placement['page']
        x, y = placement['x_points'], placement['y_points']
        if placement['heading']:
            draw_field_heading(pdf, placement['heading'], common_width)
        if label_rendering == 'raster':
            draw_raster_label(pdf, lambda small: draw_label(small, row, lines, 0, 0, common_width, args.guides), x, y, common_width, HEIGHT)
        else:
            draw_label(pdf, row, lines, x, y, common_width, args.guides)
        habitat_report = None
        if args.separate_habitat:
            habitat_report = dict(habitat_plans[index])
            habitat_report.update(page=placement['page'], x_points=x,
                                  y_points=y - habitat_gap - habitat_height)
            if label_rendering == 'raster':
                draw_raster_label(pdf, lambda small: draw_habitat_label(small, habitat_report, 0, 0, args.habitat_width * 72, args.guides), x, habitat_report['y_points'], args.habitat_width * 72, habitat_height)
            else:
                draw_habitat_label(pdf, habitat_report, x, habitat_report['y_points'], args.habitat_width * 72, args.guides)
        report.append({'excel_row': row['_excel_row'], 'catalog_number': row['Catalog_Number'],
                       'field_number': row.get('Field_Number', ''),
                       'uncertainty_omitted': bool(row.get('Uncertainty')) and not keep,
                       'macrohabitat_omitted': _layout(row, templates, common_width, keep) is None,
                       'line_text_runs': [coordinate_runs(row, text, size) for text, size in zip(lines, lines.sizes)],
                       'line_font_points': lines.sizes, 'line_baselines': lines.baselines, 'lines': lines,
                       'habitat_label': habitat_report, 'placement': placement})
    pdf.save()
    temp.replace(output)
    summary = {'label_rendering': label_rendering, 'label_raster_dpi': 1200 if label_rendering == 'raster' else None, 'barcode_rendering': BARCODE_RENDERING, 'barcode_raster_dpi': 1200 if BARCODE_RENDERING == 'raster' else None, 'labels': len(rows), 'separate_habitat': args.separate_habitat,
               'physical_labels': len(rows) * (2 if args.separate_habitat else 1),
               'habitat_label_width_inches': args.habitat_width if args.separate_habitat else None,
               'habitat_max_lines': args.habitat_max_lines if args.separate_habitat else None,
               'habitat_label_height_inches': args.habitat_height if args.separate_habitat else None,
               'habitat_label_gap_inches': args.habitat_gap if args.separate_habitat else None, 'width_inches': common_width / 72, 'height_inches': HEIGHT / 72,
               'field_headings': not args.no_field_headings, 'field_font_path': str(field_path),
               'field_font_horizontal_scale': FIELD_SCALE, 'field_left_shift_points': FIELD_LEFT_SHIFT, 'habitat_field_gap_points': HABITAT_FIELD_GAP, 'layout': LAYOUT_MODE, 'font': pdfmetrics.getFont('LabelRegular').face.familyName.decode('utf-8', errors='replace'), 'font_points': MAIN_SIZE, 'habitat_font_points': 4 if args.separate_habitat else SIZE, 'catalog_font_points': CATALOG_SIZE, 'field_font_points': FIELD_SIZE,
               'barcode_scale': barcode_scale, 'barcode_dark_symbol_inches': 0.11 * barcode_scale, 'barcode_center_y_inches': HEIGHT / 144,
               'horizontal_gap_inches': gap / 72, 'vertical_gap_inches': gap / 72, 'text_horizontal_scale': TEXT_SCALE, 'coordinate_horizontal_scale': COORDINATE_SCALE, 'barcode': 'Data Matrix',
               'macrohabitat_omitted_count': sum(r['macrohabitat_omitted'] for r in report),
               'uncertainty_omitted_count': sum(r['uncertainty_omitted'] for r in report), 'records': report}
    output.with_suffix('.json').write_text(json.dumps(summary, indent=2))
    habitat_description = '4 pt on separate habitat labels' if args.separate_habitat else '3.5 pt habitat'
    print(f"Created {output}: {len(rows)} main labels, {common_width / 72:.2f} × {HEIGHT / 72:.2f} in, {family} {MAIN_SIZE:g} pt main / {habitat_description}.")
    if args.separate_habitat:
        print(f"Added {len(rows)} habitat labels, {args.habitat_width:.2f} × {args.habitat_height:.2f} in, {args.habitat_gap:.2f}-inch gap.")
    print(f"Field Number: {FIELD_SIZE:g} pt, {field_path.name}; block headings {'off' if args.no_field_headings else 'on'}.")
    print(f"Macrohabitat omitted on {summary['macrohabitat_omitted_count']} labels.")
    print(f"Uncertainty omitted on {summary['uncertainty_omitted_count']} labels. Review {output.with_suffix('.json')}.")


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('excel')
    p.add_argument('--coordinate-scale', type=float, help='Override coordinate horizontal scale (0–1); font height is unchanged')
    p.add_argument('--barcode-scale', type=float, default=1.02, help='Scale barcode and quiet margin together (1.0–1.1); text clearance adapts')
    p.add_argument('--label-rendering', choices=['vector', 'raster'], default='vector', help='Render each complete label as a 1200-dpi image (requires Poppler); headings stay vector')
    p.add_argument('--barcode-rendering', choices=['vector', 'raster'], default='raster', help='Barcode rendering; raster embeds a 1200-dpi image with unchanged size and quiet margin')
    p.add_argument('-o', '--output', default='output/labels.pdf')
    p.add_argument('--sheet')
    p.add_argument('--layout', choices=['lines', 'flow'], default='lines', help='Separate fields into lines, or wrap comma-separated fields')
    p.add_argument('--font-family', choices=['arial', 'verdana'], default='arial', help='Main, catalog, and habitat font family (default arial)')
    p.add_argument('--italic-font', help='Matching italic .ttf path; defaults to the selected family')
    p.add_argument('--main-size', type=float, choices=[3.5, 4.0], default=4.0, help='Font size for the first five physical text lines')
    p.add_argument('--template', help='JSON list of lines with [Excel_Column] placeholders')
    p.add_argument('--font', help='Regular .ttf path; defaults to the selected family')
    p.add_argument('--width', type=float, default=.68, help='Label width in inches; must never exceed 0.68')
    p.add_argument('--max-width', type=float, default=None, help='Deprecated: must equal --width; automatic widening is disabled')
    p.add_argument('--height', type=float, default=.33, help='Label height in inches (0.33–10.5); barcode stays vertically centered')
    p.add_argument('--separate-habitat', action='store_true', help='Place a smaller 4-point habitat label (up to three lines) directly below each main label')
    p.add_argument('--habitat-height', type=float, default=.20, help='Separate habitat label height in inches (default 0.20)')
    p.add_argument('--habitat-width', type=float, default=.62, help='Separate habitat label width in inches (default 0.62; must be less than main width)')
    p.add_argument('--habitat-max-lines', type=int, choices=[2, 3], default=3, help='Maximum physical lines on separate habitat label (default 3)')
    p.add_argument('--habitat-gap', type=float, default=.02, help='Gap between main and separate habitat label in inches (default 0.02)')
    p.add_argument('--field-size', type=float, choices=[3.5, 4.0], default=4.0, help='Field-number font size (default 4)')
    p.add_argument('--field-font', help='Field-number .ttf font (prefers Verdana beside Arial, otherwise regular Arial)')
    p.add_argument('--heading-font', help='Bold .ttf font for 5-point field block headings (default selected-family Bold)')
    p.add_argument('--no-field-headings', action='store_true', help='Disable field block headings above labels')
    p.add_argument('--guides', action='store_true', help='Add light cutting rectangles')
    args = p.parse_args()
    try:
        generate(args)
    except (ValueError, KeyError, OSError) as exc:
        p.exit(1, f'Error: {exc}\n')

if __name__ == '__main__':
    main()
