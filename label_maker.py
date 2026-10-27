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


def read_rows(path, sheet=None, require_catalog=True):
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
        if require_catalog and not row.get('Catalog_Number'):
            raise ValueError(f'Excel row {number}: Catalog_Number is missing.')
        for alias, source in ALIASES.items():
            row.setdefault(alias, row.get(source, ''))
        if row.get('Collector_Label'):
            row['Collector_Label'] = re.sub(r'\b([A-Z])\. +(?=[A-Z])', r'\1.', row['Collector_Label'])
        row['_excel_row'] = number
        rows.append(row)
    book.close()
    ids = [r.get('Catalog_Number', '') for r in rows]
    if require_catalog and len(ids) != len(set(ids)):
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


def coordinate_runs(row, line, size, combined=False):
    coordinates = f"{row.get('Latitude', '')} {row.get('Longitude', '')}"
    uncertainty = f"±{row.get('Uncertainty', '')}m"
    if (LAYOUT_MODE == 'flow' or combined) and row.get('Latitude') and row.get('Longitude') and coordinates in line:
        start = line.index(coordinates)
        finish = start + len(coordinates)
        runs = [(line[:start], size, TEXT_SCALE)] if start else []
        runs.append((coordinates, size, COORDINATE_SCALE))
        if row.get('Uncertainty') and line[finish:].startswith(uncertainty):
            runs.append((uncertainty, 3.5, TEXT_SCALE))
            finish += len(uncertainty)
        if finish < len(line):
            runs.append((line[finish:], size, TEXT_SCALE))
        return runs
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
        # Coordinates (and their uncertainty) form one unbreakable unit.
        if '[Latitude]' in templates[i] and '[Longitude]' in templates[i]:
            parts = [' '.join(parts)]
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
                measured = run_width(coordinate_runs(row, text, size)) if row.get('Latitude') and row.get('Longitude') and f"{row['Latitude']} {row['Longitude']}" in text else fit_width(text, size)
                if measured <= available:
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
    draw_raster_matrix(pdf, img, bx, by)
    event = row.get('Field_Number', '')
    obj = pdf.beginText(bx + BAR - field_width(event) - FIELD_LEFT_SHIFT, y + EVENT_Y)
    obj.setFont(FIELD_FONT, FIELD_SIZE)
    obj.setCharSpace(0)
    obj.setHorizScale(FIELD_SCALE * 100)
    obj.textOut(event)
    pdf.drawText(obj)


def draw_raster_matrix(pdf, img, bx, by):
    """All Data Matrix symbols are embedded as nearest-neighbor 1200-dpi images."""
    pixels = round(BAR / 72 * 1200)
    raster = img.resize((pixels, pixels), Image.Resampling.NEAREST)
    pdf.drawImage(ImageReader(raster), bx, by, width=BAR, height=BAR)


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


def habitat_label_plan(row, label_width, height, max_lines=3, include_method=False):
    fields = [' '.join(row[key].split()) for key in ('MacroHabitat_Label', 'Microhabitat_Label')]
    field_sources = ['macrohabitat', 'microhabitat']
    if include_method:
        keys = ('Sampling_Protocol_Label', 'Sampling_Protocol', 'Method')
        if not any(key in row for key in keys):
            raise ValueError('Including method requires Sampling_Protocol_Label, Sampling_Protocol, or Method.')
        method = next((row[key] for key in keys if row.get(key)), '')
        if method:
            fields.append(' '.join(method.split()))
            field_sources.append('method')
    texts, sources = [], []
    available = label_width - 2 * PAD - .3
    for field, source in zip(fields, field_sources):
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
    micro_lines = [text for text, source in zip(texts, sources) if source == 'microhabitat']
    micro_emphasis = iter(label_italic_ranges(micro_lines, fields[1]))
    emphasis = [next(micro_emphasis) if source == 'microhabitat' else [] for source in sources]
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


def configure_dimensions(width_inches, height_inches, max_width=None, min_height=.33):
    global HEIGHT, BAR_Y
    if not math.isfinite(width_inches) or not 0 < width_inches <= MAX_WIDTH_INCHES:
        raise ValueError('Label width must be greater than 0 and never exceed 0.68 inches (four-row unit-tray limit).')
    if max_width is not None and (not math.isfinite(max_width) or max_width != width_inches):
        raise ValueError('--max-width no longer allows widening; omit it or set it equal to --width (at most 0.68 inches).')
    if not math.isfinite(height_inches) or not min_height <= height_inches <= 10.5:
        raise ValueError(f'Label height must be between {min_height:g} and 10.5 inches, within a Letter page.')
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


def special_label_plan(row, mode, label_width, show_catalog=True):
    """Plan barcode/identification labels without locality fields or event codes."""
    catalog_scale = .65
    if show_catalog and pdfmetrics.stringWidth(row['Catalog_Number'], 'LabelRegular', 4) * catalog_scale > HEIGHT - 2 * CATALOG_PAD:
        raise ValueError(f"Excel row {row['_excel_row']}: catalog number is too long for a 4-point vertical code; increase --height.")
    available = label_width - 2 * PAD - SIDE - BAR - 1.7
    if mode == 'pin-barcode':
        blank_width = label_width - PAD - SIDE - BAR
        if blank_width - PAD < .08 * 72:
            raise ValueError('Pin barcode label needs at least 0.08 inches of blank pin space left of the barcode margin; increase --width.')
        return []
    if available <= 0:
        raise ValueError('Label width leaves no room beside the barcode.')
    if mode == 'barcode':
        blocks = [('Unique Specimen Identifier', False)]
    else:
        for key in ('Genus', 'Species', 'Determiner_Label'):
            if not row.get(key):
                raise ValueError(f"Excel row {row['_excel_row']}: {key} is required for determination labels.")
        name = row['Genus'] + ' ' + row['Species']
        author = row.get('Author', '')
        full_name = name + (' ' + author if author else '')
        if author and fit_width(full_name, 4) > available:
            blocks = [(name, True), (author, False)]
        else:
            blocks = [(full_name, True)]
        determiner = row['Determiner_Label']
        if not re.match(r'^Det\.\s', determiner, re.IGNORECASE):
            determiner = 'Det. ' + determiner
        blocks.append((determiner, False))
    lines = []
    for text, scientific in blocks:
        offset = 0
        current = ''
        line_start = 0
        italic_end = (len(row['Genus']) if row['Species'] in {'sp.', 'spp.', 'indet'} else len(row['Genus'] + ' ' + row['Species'])) if scientific else 0
        for word in text.split():
            candidate = (current + ' ' + word).strip()
            if fit_width(candidate, 4) > available:
                if not current:
                    raise ValueError(f"Excel row {row['_excel_row']}: word {word!r} cannot fit; increase --width or shorten label text.")
                lines.append((current, max(0, min(len(current), italic_end-line_start))))
                offset += len(current) + 1
                line_start = offset
                current = word
                if fit_width(word, 4) > available:
                    raise ValueError(f"Excel row {row['_excel_row']}: word {word!r} cannot fit.")
            else:
                current = candidate
        if current:
            lines.append((current, max(0, min(len(current), italic_end-line_start))))
    leading = 4.5
    block_height = (len(lines)-1)*leading + 4
    if block_height > HEIGHT-2*PAD:
        raise ValueError(f"Excel row {row['_excel_row']}: text needs more height; shorten it or increase --height.")
    top_baseline = (HEIGHT+block_height)/2-3.3
    return [dict(text=t, italic_end=end, baseline=top_baseline-i*leading) for i,(t,end) in enumerate(lines)]


def draw_special_label(pdf, row, plan, x, y, label_width, guides, show_catalog):
    if guides:
        pdf.setStrokeColorRGB(.8,.8,.8); pdf.setLineWidth(.15)
        pdf.rect(x,y,label_width,HEIGHT); pdf.setStrokeColorRGB(0,0,0)
    for line in plan:
        end = line['italic_end']
        draw_text(pdf,x+PAD,y+line['baseline'],line['text'],size=4,italic_ranges=[(0,end)] if end else ())
    bx=x+label_width-PAD-SIDE-BAR
    draw_raster_matrix(pdf, matrix(row['Catalog_Number']), bx, y+BAR_Y)
    if show_catalog:
        length=pdfmetrics.stringWidth(row['Catalog_Number'],'LabelRegular',4)*.65
        pdf.saveState();pdf.translate(x+label_width-PAD-.75,y+(HEIGHT-length)/2);pdf.rotate(90)
        draw_text(pdf,0,0,row['Catalog_Number'],size=4,horizontal_scale=.65);pdf.restoreState()


def generate_special(args, rows):
    if args.separate_habitat or args.template or args.include_method:
        raise ValueError('Barcode/determination modes do not use --template, --separate-habitat, or --include-method.')
    show_catalog=not args.omit_catalog_text
    plans=[special_label_plan(r,args.mode,args.width*72,show_catalog) for r in rows]
    # Deliberately suppress Field_Number printing AND headings in these modes.
    positions=sheet_positions(rows,args.width*72,HEIGHT,False)
    output=Path(args.output);output.parent.mkdir(parents=True,exist_ok=True)
    temp=output.with_suffix('.tmp.pdf');pdf=canvas.Canvas(str(temp),pagesize=(612,792))
    pdf.setTitle('Pittacia Insectorum — '+args.mode+' labels')
    page=1
    for row,plan,pos in zip(rows,plans,positions):
        if pos['page']!=page:pdf.showPage();page=pos['page']
        x,y=pos['x_points'],pos['y_points']
        if args.label_rendering=='raster':
            draw_raster_label(pdf,lambda small:draw_special_label(small,row,plan,0,0,args.width*72,args.guides,show_catalog),x,y,args.width*72,HEIGHT)
        else:draw_special_label(pdf,row,plan,x,y,args.width*72,args.guides,show_catalog)
    pdf.save();temp.replace(output)
    report=dict(pin_blank_width_inches=(args.width*72-PAD-SIDE-BAR)/72 if args.mode == 'pin-barcode' else None, mode=args.mode,labels=len(rows),width_inches=args.width,height_inches=args.height,
                barcode_scale=args.barcode_scale,barcode_rendering=BARCODE_RENDERING,
                label_rendering=args.label_rendering,barcode_raster_dpi=1200 if BARCODE_RENDERING=='raster' else None,
                catalog_font_points=4,catalog_horizontal_scale=.65,catalog_text=show_catalog,
                field_headings=False,barcode_center_y_inches=HEIGHT/144,
                records=[dict(catalog_number=r['Catalog_Number'],excel_row=r['_excel_row'],lines=plan,placement=pos) for r,plan,pos in zip(rows,plans,positions)])
    output.with_suffix('.json').write_text(json.dumps(report,indent=2))
    print(f'Created {output}: {len(rows)} {args.mode} labels, {args.width:g} × {args.height:g} in; 4-point text; no collection-event codes.')


def first_collector_label(row):
    if row.get('First_Collector_Label'):
        return row['First_Collector_Label'].strip()
    text = row.get('Collector_Label', '').strip()
    text = re.split(r';|\s+(?:and|&)\s+', text, maxsplit=1)[0].strip()
    # Initial-first display names use commas between collectors. Surname-first
    # names may contain a comma within one person's name; preserve that form.
    if re.match(r'^[A-Z]\.', text):
        text = text.split(',', 1)[0].strip()
    elif ',' in text and not re.match(r'^[^,]+,\s*(?:[A-Z]\.\s*)+$', text):
        text = text.split(',', 1)[0].strip()
    return text


def collection_only_runs(row, text):
    # All printed characters, including coordinate uncertainty, stay at 4 pt.
    return [(part, 4, scale) for part, _, scale in coordinate_runs(row, text, 4, combined=True)]


def collection_only_plan(row, label_width, templates):
    """Five stable field lines, 4-point type, first collector only."""
    display = dict(row, Collector_Label=first_collector_label(row))
    event = ' #' + row['Field_Number'] if row.get('Field_Number') else ''
    if templates in (DEFAULT_LINES, DEFAULT_LINES[:5]):
        fixed = [DEFAULT_LINES[0], DEFAULT_LINES[1],
                 '[Latitude] [Longitude]±[Uncertainty]m, [Verbatim_Elevation]',
                 '[Verbatim_Date], [Collector_Label]' + event,
                 HABITAT_TEMPLATE if templates == DEFAULT_LINES else '']
    else:
        if len(templates) > 5:
            raise ValueError('Collection-only custom templates must have at most five fixed lines.')
        fixed = [t + (event if '[Collector_Label]' in t else '') for t in templates]
    texts = render_lines(display, fixed, True)
    if not any(text.strip() for text in texts):
        raise ValueError(f"Excel row {row['_excel_row']}: no collection text to print.")
    available = label_width - 2 * PAD
    for index, text in enumerate(texts, 1):
        measured = sum(fit_width(part, size) / TEXT_SCALE * scale for part, size, scale in collection_only_runs(display,text))
        if measured > available:
            raise ValueError(f"Excel row {row['_excel_row']}: collection-only line {index} cannot fit at 4 pt in five fixed lines. Shorten its label fields or increase --width (maximum 0.68); text is not shrunk or moved between lines.")
    sizes = [4] * len(texts)
    emphasis = label_italic_ranges(texts, row.get('Microhabitat_Label', ''))
    # A fixed grid keeps identical line positions even when habitat is separate.
    leading = 4.4
    above = max(pdfmetrics.getAscent(name)*4/1000 for name in ('LabelRegular','LabelItalic'))
    below = max(-pdfmetrics.getDescent(name)*4/1000 for name in ('LabelRegular','LabelItalic'))
    block_height = above + leading*(len(texts)-1) + below
    if block_height > HEIGHT-2*PAD:
        raise ValueError(f"Excel row {row['_excel_row']}: five 4-point lines need more height; increase --height.")
    first=(HEIGHT+block_height)/2-above
    return dict(lines=texts,line_font_points=sizes,baselines=[first-i*leading for i in range(len(texts))],italic_ranges=emphasis,first_collector_label=display['Collector_Label'])


def generate_collection_only(args, rows):
    if args.main_size != 4:
        raise ValueError('Collection-only labels use fixed 4-point text; omit --main-size or set it to 4.')
    if args.omit_catalog_text:
        raise ValueError('Collection-only labels never print catalog numbers; omit --omit-catalog-text.')
    templates = json.loads(Path(args.template).read_text()) if args.template else DEFAULT_LINES
    if not isinstance(templates, list) or not templates or not all(isinstance(t,str) for t in templates):
        raise ValueError('Template must be a nonempty list of text entries.')
    if any(key in text for text in templates for key in ('[Catalog_Number]', '[Field_Number]')):
        raise ValueError('Collection-only templates must not include catalog/field number placeholders; Field_Number is appended automatically.')
    if args.separate_habitat:
        if not math.isfinite(args.habitat_width) or not 0 < args.habitat_width < args.width:
            raise ValueError('Separate habitat width must be smaller than the main label width.')
        if not math.isfinite(args.habitat_height) or args.habitat_height <= 0 or not math.isfinite(args.habitat_gap) or args.habitat_gap < 0:
            raise ValueError('Habitat height must be positive and gap nonnegative.')
        templates=separate_habitat_templates(templates)
    width=args.width*72
    plans=[collection_only_plan(row,width,templates) for row in rows]
    habitats=[habitat_label_plan(row,args.habitat_width*72,args.habitat_height*72,args.habitat_max_lines,args.include_method) for row in rows] if args.separate_habitat else []
    pair_height=HEIGHT+(args.habitat_gap+args.habitat_height)*72 if args.separate_habitat else HEIGHT
    positions=sheet_positions(rows,width,pair_height,not args.no_field_headings)
    output=Path(args.output);output.parent.mkdir(parents=True,exist_ok=True)
    temp=output.with_suffix('.tmp.pdf');pdf=canvas.Canvas(str(temp),pagesize=(612,792));pdf.setTitle('Pittacia Insectorum — collection-only labels')
    def draw(pdf,row,plan,x,y):
        if args.guides:
            pdf.setStrokeColorRGB(.8,.8,.8);pdf.setLineWidth(.15);pdf.rect(x,y,width,HEIGHT);pdf.setStrokeColorRGB(0,0,0)
        for text,size,baseline,spans in zip(plan['lines'],plan['line_font_points'],plan['baselines'],plan['italic_ranges']):
            cursor=x+PAD;offset=0
            for part,part_size,scale in collection_only_runs(row,text):
                local_spans=[(max(a,offset)-offset,min(b,offset+len(part))-offset) for a,b in spans if a<offset+len(part) and b>offset]
                draw_text(pdf,cursor,y+baseline,part,size=part_size,italic_ranges=local_spans,horizontal_scale=scale)
                cursor+=pdfmetrics.stringWidth(part,'LabelRegular',part_size)*scale
                offset+=len(part)
    page=1;records=[]
    for i,(row,plan,pos) in enumerate(zip(rows,plans,positions)):
        if pos['page']!=page:pdf.showPage();page=pos['page']
        x,y=pos['x_points'],pos['y_points']
        if pos['heading']:draw_field_heading(pdf,pos['heading'],width)
        if args.label_rendering=='raster':draw_raster_label(pdf,lambda small:draw(small,row,plan,0,0),x,y,width,HEIGHT)
        else:draw(pdf,row,plan,x,y)
        habitat=None
        if args.separate_habitat:
            habitat=dict(habitats[i],page=pos['page'],x_points=x,y_points=y-args.habitat_gap*72-args.habitat_height*72)
            if args.label_rendering=='raster':draw_raster_label(pdf,lambda small:draw_habitat_label(small,habitat,0,0,args.habitat_width*72,args.guides),x,habitat['y_points'],args.habitat_width*72,args.habitat_height*72)
            else:draw_habitat_label(pdf,habitat,x,habitat['y_points'],args.habitat_width*72,args.guides)
        records.append(dict(excel_row=row['_excel_row'],field_number=row.get('Field_Number',''),catalog_number=None,placement=pos,habitat_label=habitat,**plan))
    pdf.save();temp.replace(output)
    report=dict(mode='collection-only',barcode=None,barcode_rendering=None,barcode_raster_dpi=None,catalog_text=False,font=pdfmetrics.getFont('LabelRegular').face.familyName.decode(),labels=len(rows),physical_labels=len(rows)*(2 if args.separate_habitat else 1),width_inches=args.width,height_inches=args.height,max_main_lines=5,text_horizontal_scale=TEXT_SCALE,main_font_points=4,text_width_points=width-2*PAD,separate_habitat=args.separate_habitat,habitat_label_width_inches=args.habitat_width if args.separate_habitat else None,habitat_label_height_inches=args.habitat_height if args.separate_habitat else None,habitat_label_gap_inches=args.habitat_gap if args.separate_habitat else None,include_method=args.include_method,records=records)
    output.with_suffix('.json').write_text(json.dumps(report,indent=2))
    print(f'Created {output}: {len(rows)} collection-only labels, {args.width:g} × {args.height:g} inches; five fixed 4-point lines; first collector only; no barcode or catalog text.')


def generate(args):
    global BARCODE_RENDERING, BAR
    mode = getattr(args, 'mode', 'locality')
    if args.width is None: args.width = .32 if mode == 'pin-barcode' else .55 if mode == 'barcode' else .68
    if args.height is None: args.height = .30 if mode in {'barcode','pin-barcode'} else .33
    if args.habitat_width is None: args.habitat_width = .56 if mode == 'collection-only' else .62
    if args.barcode_scale is None: args.barcode_scale = 1.05 if mode in {'barcode','determination'} else 1.02
    barcode_scale = args.barcode_scale
    if not math.isfinite(barcode_scale) or not 1 <= barcode_scale <= 1.1:
        raise ValueError('--barcode-scale must be between 1.0 and 1.1.')
    BAR = 7.92 * 18 / 14 * barcode_scale
    if getattr(args, 'barcode_rendering', 'raster') != 'raster':
        raise ValueError('Data Matrix barcodes must use 1200-dpi raster rendering.')
    BARCODE_RENDERING = 'raster'
    label_rendering = getattr(args, 'label_rendering', 'vector')
    if label_rendering == 'raster' and not shutil.which('pdftoppm'):
        raise ValueError('Whole-label raster rendering requires pdftoppm (Poppler). On macOS: brew install poppler.')
    global MAIN_SIZE, FONT_PATH, LAYOUT_MODE, ITALIC_FONT_PATH, FIELD_SIZE, FIELD_FONT, TEXT_SCALE, COORDINATE_SCALE
    configure_dimensions(args.width, args.height, args.max_width, min_height=.20 if mode != 'locality' else .33)
    FIELD_SIZE = args.field_size
    MAIN_SIZE = args.main_size
    LAYOUT_MODE = args.layout
    family = 'Verdana' if args.font_family == 'verdana' else 'Arial'
    TEXT_SCALE = (.62 if family == 'Verdana' else .70) if mode == 'collection-only' else (.70 if family == 'Verdana' else .75)
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
    if not heading_path.is_file() and not args.no_field_headings and mode in {'locality','collection-only'}:
        raise ValueError('Bold heading font not found. Supply --heading-font with a matching bold .ttf path.')
    if heading_path.is_file():
        pdfmetrics.registerFont(TTFont('FieldHeading', str(heading_path)))
    if mode == 'locality' and getattr(args, 'omit_catalog_text', False):
        raise ValueError('--omit-catalog-text is only available for barcode/determination labels.')
    if args.include_method and not args.separate_habitat:
        raise ValueError('--include-method requires --separate-habitat.')
    rows = read_rows(args.excel, args.sheet, require_catalog=mode != 'collection-only')
    if mode == 'collection-only':
        generate_collection_only(args, rows)
        return
    if mode != 'locality':
        generate_special(args, rows)
        return
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
    habitat_plans = [habitat_label_plan(row, args.habitat_width * 72, habitat_height, args.habitat_max_lines, include_method=args.include_method) for row in rows] if args.separate_habitat else []
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
    summary = {'include_method': args.include_method, 'label_rendering': label_rendering, 'label_raster_dpi': 1200 if label_rendering == 'raster' else None, 'barcode_rendering': BARCODE_RENDERING, 'barcode_raster_dpi': 1200 if BARCODE_RENDERING == 'raster' else None, 'labels': len(rows), 'separate_habitat': args.separate_habitat,
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
    p.add_argument('--mode', choices=['locality','barcode','determination','pin-barcode','collection-only'], default='locality', help='Locality labels, compact generic barcode labels, or taxon/determiner labels')
    p.add_argument('--include-method', action='store_true', help='Add sampling protocol after microhabitat on separate habitat labels; three lines total')
    p.add_argument('--omit-catalog-text', action='store_true', help='Barcode/determination modes: encode the catalog number without printing it as text')
    p.add_argument('--coordinate-scale', type=float, help='Override coordinate horizontal scale (0–1); font height is unchanged')
    p.add_argument('--barcode-scale', type=float, default=None, help='Scale barcode and quiet margin together (1.0–1.1); text clearance adapts')
    p.add_argument('--label-rendering', choices=['vector', 'raster'], default='vector', help='Render each complete label as a 1200-dpi image (requires Poppler); headings stay vector')
    p.add_argument('--barcode-rendering', choices=['raster'], default='raster', help='Compatibility option: all barcodes use 1200-dpi raster images')
    p.add_argument('-o', '--output', default='output/labels.pdf')
    p.add_argument('--sheet')
    p.add_argument('--layout', choices=['lines', 'flow'], default='lines', help='Separate fields into lines, or wrap comma-separated fields')
    p.add_argument('--font-family', choices=['arial', 'verdana'], default='verdana', help='Main, catalog, and habitat font family (default verdana)')
    p.add_argument('--italic-font', help='Matching italic .ttf path; defaults to the selected family')
    p.add_argument('--main-size', type=float, choices=[3.5, 4.0], default=4.0, help='Font size for the first five physical text lines')
    p.add_argument('--template', help='JSON list of lines with [Excel_Column] placeholders')
    p.add_argument('--font', help='Regular .ttf path; defaults to the selected family')
    p.add_argument('--width', type=float, default=None, help='Label width in inches; must never exceed 0.68')
    p.add_argument('--max-width', type=float, default=None, help='Deprecated: must equal --width; automatic widening is disabled')
    p.add_argument('--height', type=float, default=None, help='Label height in inches (0.33–10.5); barcode stays vertically centered')
    p.add_argument('--separate-habitat', action='store_true', help='Place a smaller 4-point habitat label (up to three lines) directly below each main label')
    p.add_argument('--habitat-height', type=float, default=.20, help='Separate habitat label height in inches (default 0.20)')
    p.add_argument('--habitat-width', type=float, default=None, help='Separate habitat label width in inches (default 0.62; must be less than main width)')
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
