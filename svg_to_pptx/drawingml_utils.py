"""Coordinate helpers, color parsing, and font utilities for DrawingML conversion."""

from __future__ import annotations

import math
import re
from xml.etree import ElementTree as ET

from .drawingml_context import ConvertContext

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SVG_NS = 'http://www.w3.org/2000/svg'
XLINK_NS = 'http://www.w3.org/1999/xlink'

EMU_PER_PX = 9525  # 1 SVG px = 9525 EMU (96 DPI)
FONT_PX_TO_HUNDREDTHS_PT = 75  # 1px = 0.75pt -> 75 hundredths-of-a-point
ANGLE_UNIT = 60000  # DrawingML angle: 60000ths of a degree

# SVG attributes inheritable from parent <g>
INHERITABLE_ATTRS = [
    'fill', 'stroke', 'stroke-width', 'stroke-dasharray', 'stroke-linecap',
    'stroke-linejoin', 'opacity', 'fill-opacity', 'stroke-opacity',
    'font-family', 'font-size', 'font-weight', 'font-style',
    'text-anchor', 'letter-spacing', 'text-decoration',
]

# Known East Asian fonts
EA_FONTS = {
    'PingFang SC', 'PingFang TC', 'PingFang HK',
    'Microsoft YaHei', 'Microsoft JhengHei',
    'SimSun', 'SimHei', 'FangSong', 'KaiTi', 'STKaiti',
    'STHeiti', 'STSong', 'STFangsong', 'STXihei', 'STZhongsong',
    'Hiragino Sans', 'Hiragino Sans GB', 'Hiragino Mincho ProN',
    'Noto Sans SC', 'Noto Sans TC', 'Noto Serif SC', 'Noto Serif TC',
    'Source Han Sans SC', 'Source Han Sans TC',
    'Source Han Serif SC', 'Source Han Serif TC',
    'WenQuanYi Micro Hei', 'WenQuanYi Zen Hei',
    'YouYuan', 'LiSu', 'HuaWenKaiTi',
    'Songti SC', 'Songti TC',
}
SYSTEM_FONTS = {'system-ui', '-apple-system', 'BlinkMacSystemFont'}

# macOS/Linux-only fonts -> Windows equivalents
FONT_FALLBACK_WIN = {
    'PingFang SC': 'Microsoft YaHei',
    'PingFang TC': 'Microsoft JhengHei',
    'PingFang HK': 'Microsoft JhengHei',
    'Hiragino Sans': 'Microsoft YaHei',
    'Hiragino Sans GB': 'Microsoft YaHei',
    'Hiragino Mincho ProN': 'SimSun',
    'STHeiti': 'SimHei',
    'STSong': 'SimSun',
    'STKaiti': 'KaiTi',
    'STFangsong': 'FangSong',
    'STXihei': 'Microsoft YaHei',
    'STZhongsong': 'SimSun',
    'Songti SC': 'SimSun',
    'Songti TC': 'SimSun',
    'Noto Sans SC': 'Microsoft YaHei',
    'Noto Sans TC': 'Microsoft JhengHei',
    'Noto Serif SC': 'SimSun',
    'Noto Serif TC': 'SimSun',
    'Source Han Sans SC': 'Microsoft YaHei',
    'Source Han Sans TC': 'Microsoft JhengHei',
    'Source Han Serif SC': 'SimSun',
    'Source Han Serif TC': 'SimSun',
    'WenQuanYi Micro Hei': 'Microsoft YaHei',
    'WenQuanYi Zen Hei': 'Microsoft YaHei',
    # Latin fonts (macOS / Linux / Web -> Windows)
    'SF Pro': 'Segoe UI',
    'SF Pro Display': 'Segoe UI',
    'SF Pro Text': 'Segoe UI',
    'SF Mono': 'Consolas',
    'Menlo': 'Consolas',
    'Monaco': 'Consolas',
    'Helvetica Neue': 'Arial',
    'Helvetica': 'Arial',
    'Roboto': 'Segoe UI',
    'Ubuntu': 'Segoe UI',
    'Liberation Sans': 'Arial',
    'Liberation Serif': 'Times New Roman',
    'Liberation Mono': 'Consolas',
    'DejaVu Sans': 'Segoe UI',
    'DejaVu Serif': 'Times New Roman',
    'DejaVu Sans Mono': 'Consolas',
}

GENERIC_FONT_MAP = {
    'monospace': 'Consolas',
    'sans-serif': 'Segoe UI',
    'serif': 'Times New Roman',
}

COLOR_KEYWORDS = {
    'black': '000000',
    'white': 'FFFFFF',
    'red': 'FF0000',
    'green': '008000',
    'blue': '0000FF',
    'yellow': 'FFFF00',
    'gray': '808080',
    'grey': '808080',
    'orange': 'FFA500',
    'purple': '800080',
    'pink': 'FFC0CB',
    'brown': 'A52A2A',
}

_NUMBER_RE = re.compile(r'[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?')

# When the latin font is serif and no EA font is specified,
# prefer SimSun (serif CJK) over Microsoft YaHei (sans-serif CJK).
_SERIF_LATIN = {
    'Times New Roman', 'Georgia', 'Garamond', 'Palatino', 'Palatino Linotype',
    'Book Antiqua', 'Cambria', 'SimSun', 'Liberation Serif', 'DejaVu Serif',
}

# SVG stroke-dasharray -> DrawingML prstDash
DASH_PRESETS = {
    '4,4': 'dash',  '4 4': 'dash',
    '6,3': 'dash',  '6 3': 'dash',
    '2,2': 'sysDot', '2 2': 'sysDot',
    '8,4': 'lgDash', '8 4': 'lgDash',
    '8,4,2,4': 'lgDashDot', '8 4 2 4': 'lgDashDot',
}


# ---------------------------------------------------------------------------
# Coordinate helpers
# ---------------------------------------------------------------------------

def px_to_emu(px: float) -> int:
    """Convert SVG pixels to EMU."""
    return round(px * EMU_PER_PX)


def _f(val: str | None, default: float = 0.0) -> float:
    """Parse a float attribute value, returning default if missing."""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)
    text = str(val).strip()
    if not text:
        return default
    try:
        return float(text)
    except (ValueError, TypeError):
        pass

    match = _NUMBER_RE.match(text)
    if not match:
        return default

    try:
        number = float(match.group(0))
    except ValueError:
        return default

    if text.endswith('%'):
        return number / 100.0
    return number


def _parse_style_declarations(style_attr: str) -> dict[str, str]:
    """Parse a CSS declaration string into a property map."""
    styles: dict[str, str] = {}
    if not style_attr:
        return styles

    for part in style_attr.split(';'):
        if ':' not in part:
            continue
        key, val = part.split(':', 1)
        key = key.strip()
        val = val.strip()
        if key:
            styles[key] = val
    return styles


def _should_fallback_to_style(attr: str, value: str) -> bool:
    """Whether an explicit SVG attribute should defer to inline/class style."""
    stripped = value.strip()
    return attr in ('fill', 'stroke', 'color') and stripped.startswith('var(')


def _extract_inheritable_styles(
    elem: ET.Element,
    ctx: ConvertContext | None = None,
) -> dict[str, str]:
    """Extract all SVG-inheritable presentation attributes from an element."""
    styles: dict[str, str] = {}

    # Inline style has higher priority than class style but lower than
    # explicit presentation attributes.
    inline_styles = _parse_style_declarations(elem.get('style', ''))

    class_styles: dict[str, str] = {}
    class_attr = elem.get('class', '')
    if class_attr and ctx:
        for cls in class_attr.split():
            cls_map = ctx.class_styles.get(cls)
            if cls_map:
                class_styles.update(cls_map)

    for attr in INHERITABLE_ATTRS:
        val = elem.get(attr)
        if val is not None and not _should_fallback_to_style(attr, val):
            styles[attr] = val
            continue
        if attr in inline_styles:
            styles[attr] = inline_styles[attr]
            continue
        if attr in class_styles:
            styles[attr] = class_styles[attr]
            continue
        if val is not None:
            styles[attr] = val
    return styles


def _get_attr(elem: ET.Element, attr: str, ctx: ConvertContext) -> str | None:
    """Get effective attribute: element's own value first, then inherited."""
    val = elem.get(attr)
    if val is not None and not _should_fallback_to_style(attr, val):
        return val

    # style="a:b;c:d"
    style_map = _parse_style_declarations(elem.get('style', ''))
    if attr in style_map:
        return style_map[attr]

    # class="foo bar"
    class_attr = elem.get('class', '')
    if class_attr:
        for cls in class_attr.split():
            cls_map = ctx.class_styles.get(cls)
            if cls_map and attr in cls_map:
                return cls_map[attr]

    if val is not None:
        return val

    return ctx.inherited_styles.get(attr)


def ctx_x(val: float, ctx: ConvertContext) -> float:
    """Apply context scale + translate to an X coordinate."""
    return val * ctx.scale_x + ctx.translate_x


def ctx_y(val: float, ctx: ConvertContext) -> float:
    """Apply context scale + translate to a Y coordinate."""
    return val * ctx.scale_y + ctx.translate_y


def ctx_w(val: float, ctx: ConvertContext) -> float:
    """Apply context scale to a width value."""
    return val * ctx.scale_x


def ctx_h(val: float, ctx: ConvertContext) -> float:
    """Apply context scale to a height value."""
    return val * ctx.scale_y


def multiply_svg_matrices(
    left: tuple[float, float, float, float, float, float],
    right: tuple[float, float, float, float, float, float],
) -> tuple[float, float, float, float, float, float]:
    """Multiply two SVG affine matrices."""
    a1, b1, c1, d1, e1, f1 = left
    a2, b2, c2, d2, e2, f2 = right
    return (
        a1 * a2 + c1 * b2,
        b1 * a2 + d1 * b2,
        a1 * c2 + c1 * d2,
        b1 * c2 + d1 * d2,
        a1 * e2 + c1 * f2 + e1,
        b1 * e2 + d1 * f2 + f1,
    )


def parse_transform_matrix(transform_str: str) -> tuple[float, float, float, float, float, float]:
    """Parse an SVG transform string into an affine matrix.

    Returns (a, b, c, d, e, f) for:
      [a c e]
      [b d f]
      [0 0 1]
    """
    if not transform_str:
        return (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

    matrix = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    for name, raw_args in re.findall(r'([A-Za-z]+)\(([^)]*)\)', transform_str):
        args = [
            float(token)
            for token in re.findall(r'[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?', raw_args)
        ]
        op = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

        if name == 'matrix' and len(args) == 6:
            op = tuple(args)  # type: ignore[assignment]
        elif name == 'translate' and args:
            tx = args[0]
            ty = args[1] if len(args) > 1 else 0.0
            op = (1.0, 0.0, 0.0, 1.0, tx, ty)
        elif name == 'scale' and args:
            sx = args[0]
            sy = args[1] if len(args) > 1 else sx
            op = (sx, 0.0, 0.0, sy, 0.0, 0.0)
        elif name == 'rotate' and args:
            angle = math.radians(args[0])
            cos_a = math.cos(angle)
            sin_a = math.sin(angle)
            rot = (cos_a, sin_a, -sin_a, cos_a, 0.0, 0.0)
            if len(args) >= 3:
                cx = args[1]
                cy = args[2]
                op = multiply_svg_matrices(
                    multiply_svg_matrices((1.0, 0.0, 0.0, 1.0, cx, cy), rot),
                    (1.0, 0.0, 0.0, 1.0, -cx, -cy),
                )
            else:
                op = rot
        elif name == 'skewX' and args:
            op = (1.0, 0.0, math.tan(math.radians(args[0])), 1.0, 0.0, 0.0)
        elif name == 'skewY' and args:
            op = (1.0, math.tan(math.radians(args[0])), 0.0, 1.0, 0.0, 0.0)

        matrix = multiply_svg_matrices(matrix, op)

    return matrix


def parse_transform_components(transform_str: str) -> tuple[float, float, float, float]:
    """Extract axis-aligned translate/scale components from an SVG transform."""
    a, b, c, d, e, f = parse_transform_matrix(transform_str)
    sx = math.copysign(math.hypot(a, b), a if abs(a) >= abs(b) else b or 1.0)
    sy = math.copysign(math.hypot(c, d), d if abs(d) >= abs(c) else c or 1.0)
    if abs(b) < 1e-8 and abs(c) < 1e-8:
        sx = a if abs(a) > 1e-8 else sx
        sy = d if abs(d) > 1e-8 else sy
    return e, f, sx or 1.0, sy or 1.0


def extract_transform_rotation_deg(transform_str: str) -> float:
    """Extract net rotation in degrees from an SVG transform."""
    if not transform_str:
        return 0.0
    a, b, _, _, _, _ = parse_transform_matrix(transform_str)
    return math.degrees(math.atan2(b, a))


# ---------------------------------------------------------------------------
# Color / style parsing
# ---------------------------------------------------------------------------

def _parse_rgb_channel(channel: str) -> int | None:
    """Parse a single CSS rgb channel into an 8-bit integer."""
    channel = channel.strip()
    if not channel:
        return None
    try:
        if channel.endswith('%'):
            return max(0, min(255, round(float(channel[:-1]) * 2.55)))
        return max(0, min(255, round(float(channel))))
    except ValueError:
        return None


def _parse_alpha_channel(alpha: str) -> float | None:
    """Parse a CSS alpha value."""
    alpha = alpha.strip()
    if not alpha:
        return None
    try:
        if alpha.endswith('%'):
            return max(0.0, min(1.0, float(alpha[:-1]) / 100.0))
        return max(0.0, min(1.0, float(alpha)))
    except ValueError:
        return None


def parse_color_value(color_str: str) -> tuple[str | None, float | None]:
    """Parse an SVG/CSS color into ('RRGGBB', alpha)."""
    if not color_str:
        return None, None

    color_str = color_str.strip()
    lower = color_str.lower()

    if lower == 'none':
        return None, None
    if lower == 'transparent':
        return '000000', 0.0

    if color_str.startswith('#'):
        hex_color = color_str[1:]
        if len(hex_color) == 3:
            hex_color = ''.join(c * 2 for c in hex_color)
        if len(hex_color) == 6 and all(c in '0123456789abcdefABCDEF' for c in hex_color):
            return hex_color.upper(), None
        return None, None

    keyword = COLOR_KEYWORDS.get(lower)
    if keyword:
        return keyword, None

    rgb_match = re.fullmatch(r'rgba?\(([^)]+)\)', color_str, flags=re.IGNORECASE)
    if not rgb_match:
        return None, None

    parts = [
        part for part in re.split(r'\s*,\s*|\s+', rgb_match.group(1).strip())
        if part and part != '/'
    ]
    if len(parts) not in (3, 4):
        return None, None

    rgb = [_parse_rgb_channel(part) for part in parts[:3]]
    if any(part is None for part in rgb):
        return None, None

    alpha = _parse_alpha_channel(parts[3]) if len(parts) == 4 else None
    return ''.join(f'{int(part):02X}' for part in rgb if part is not None), alpha


def parse_hex_color(color_str: str) -> str | None:
    """Parse common SVG/CSS color formats to 'RRGGBB'."""
    color, _ = parse_color_value(color_str)
    return color


def combine_opacity(base: float | None, extra: float | None) -> float | None:
    """Combine two opacity values multiplicatively."""
    if base is None:
        return extra
    if extra is None:
        return base
    return base * extra


def normalize_dasharray(dasharray_str: str) -> str:
    """Normalize CSS dasharray values like '4px, 3px' to '4 3'."""
    if not dasharray_str:
        return ''

    values: list[str] = []
    for part in re.split(r'[\s,]+', dasharray_str.strip()):
        if not part:
            continue
        parsed = _f(part, math.nan)
        if math.isnan(parsed):
            return dasharray_str.strip()
        if abs(parsed - round(parsed)) < 1e-8:
            values.append(str(int(round(parsed))))
        else:
            values.append(f'{parsed:.3f}'.rstrip('0').rstrip('.'))
    return ' '.join(values)


def parse_stop_style(style_str: str) -> tuple[str | None, float]:
    """Parse a gradient stop's style attribute.

    Args:
        style_str: Style string like 'stop-color:#XXX;stop-opacity:N'.

    Returns:
        (color, opacity) tuple.
    """
    color = None
    opacity = 1.0
    if not style_str:
        return color, opacity

    for part in style_str.split(';'):
        part = part.strip()
        if part.startswith('stop-color:'):
            color = parse_hex_color(part.split(':', 1)[1].strip())
        elif part.startswith('stop-opacity:'):
            try:
                opacity = float(part.split(':', 1)[1].strip())
            except ValueError:
                pass

    return color, opacity


def resolve_url_id(url_str: str) -> str | None:
    """Extract ID from 'url(#someId)' reference."""
    if not url_str:
        return None
    m = re.match(r'url\(#([^)]+)\)', url_str.strip())
    return m.group(1) if m else None


def get_effective_filter_id(elem: ET.Element, ctx: ConvertContext) -> str | None:
    """Get the effective filter ID for an element, including inherited context."""
    filt = elem.get('filter')
    if filt:
        return resolve_url_id(filt)
    return ctx.filter_id


# ---------------------------------------------------------------------------
# Font parsing
# ---------------------------------------------------------------------------

def parse_font_family(font_family_str: str) -> dict[str, str]:
    """Parse CSS font-family into latin/ea typeface names.

    Prioritizes Windows-available fonts since PPTX is primarily opened on
    Windows. macOS/Linux-only fonts are mapped via FONT_FALLBACK_WIN.
    """
    if not font_family_str:
        return {'latin': 'Segoe UI', 'ea': 'Microsoft YaHei'}

    fonts = [f.strip().strip("'\"") for f in font_family_str.split(',')]
    latin_font = None
    ea_font = None

    for font in fonts:
        if font in SYSTEM_FONTS:
            continue
        if font in GENERIC_FONT_MAP:
            resolved = GENERIC_FONT_MAP[font]
            latin_font = latin_font or resolved
            continue

        win_font = FONT_FALLBACK_WIN.get(font, font)
        if font in EA_FONTS:
            ea_font = ea_font or win_font
        else:
            latin_font = latin_font or win_font

    # PPT renders CJK text via latin typeface when ea doesn't match
    if not latin_font and ea_font:
        latin_font = ea_font

    final_latin = latin_font or 'Segoe UI'

    # EA must always be a CJK-capable font
    if not ea_font:
        ea_font = 'SimSun' if final_latin in _SERIF_LATIN else 'Microsoft YaHei'

    return {'latin': final_latin, 'ea': ea_font}


def is_cjk_char(ch: str) -> bool:
    """Check if a character is CJK (Chinese/Japanese/Korean)."""
    cp = ord(ch)
    return (0x4E00 <= cp <= 0x9FFF or 0x3400 <= cp <= 0x4DBF or
            0x2E80 <= cp <= 0x2EFF or 0x3000 <= cp <= 0x303F or
            0xFF00 <= cp <= 0xFFEF or 0xF900 <= cp <= 0xFAFF or
            0x20000 <= cp <= 0x2A6DF)


def estimate_text_width(text: str, font_size: float, font_weight: str = '400') -> float:
    """Estimate text width in SVG pixels."""
    width = 0.0
    for ch in text:
        if is_cjk_char(ch):
            width += font_size
        elif ch == ' ':
            width += font_size * 0.3
        elif ch in 'mMwWOQ':
            width += font_size * 0.75
        elif ch in 'iIlj1!|':
            width += font_size * 0.3
        else:
            width += font_size * 0.55

    if font_weight in ('bold', '600', '700', '800', '900'):
        width *= 1.05

    return width


def _xml_escape(text: str) -> str:
    """Escape XML special characters."""
    return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;'))
