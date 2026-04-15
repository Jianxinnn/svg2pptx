from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from svg_to_pptx.drawingml_converter import convert_svg_to_slide_shapes
from svg_to_pptx.drawingml_utils import EMU_PER_PX


SVG_TEMPLATE = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300">
{body}
</svg>'''


def _render_svg(body: str) -> str:
    with tempfile.TemporaryDirectory() as tmp_dir:
        svg_path = Path(tmp_dir) / 'sample.svg'
        svg_path.write_text(SVG_TEMPLATE.format(body=body), encoding='utf-8')
        slide_xml, _, _ = convert_svg_to_slide_shapes(svg_path)
    return slide_xml


class SvgRegressionTests(unittest.TestCase):
    def test_nested_scale_then_translate_group_keeps_scaled_translation(self) -> None:
        slide_xml = _render_svg(
            '<g transform="scale(2)"><g transform="translate(10,20)">'
            '<rect x="5" y="6" width="7" height="8" fill="#ff0000"/>'
            '</g></g>'
        )

        self.assertIn(f'<a:off x="{30 * EMU_PER_PX}" y="{52 * EMU_PER_PX}"/>', slide_xml)
        self.assertIn(f'<a:ext cx="{14 * EMU_PER_PX}" cy="{16 * EMU_PER_PX}"/>', slide_xml)

    def test_rect_matrix_translation_is_applied(self) -> None:
        slide_xml = _render_svg(
            '<rect x="5" y="6" width="7" height="8" fill="#00ff00" '
            'transform="matrix(1 0 0 1 10 20)"/>'
        )

        self.assertIn(f'<a:off x="{15 * EMU_PER_PX}" y="{26 * EMU_PER_PX}"/>', slide_xml)
        self.assertIn(f'<a:ext cx="{7 * EMU_PER_PX}" cy="{8 * EMU_PER_PX}"/>', slide_xml)

    def test_positioned_tspans_become_separate_text_boxes(self) -> None:
        slide_xml = _render_svg(
            '<text x="100" y="50" font-size="20" text-anchor="middle">'
            '<tspan x="100" y="50">Hello</tspan>'
            '<tspan x="100" y="80">World</tspan>'
            '</text>'
        )

        self.assertEqual(slide_xml.count('name="TextBox '), 2)
        self.assertIn('<a:t>Hello</a:t>', slide_xml)
        self.assertIn('<a:t>World</a:t>', slide_xml)

    def test_use_expands_referenced_hexagons(self) -> None:
        slide_xml = _render_svg(
            '<defs>'
            '<symbol id="hex">'
            '<polygon points="10,0 20,6 20,18 10,24 0,18 0,6" fill="#89B2D6" stroke="#56789A"/>'
            '</symbol>'
            '</defs>'
            '<use href="#hex" x="10" y="20"/>'
            '<use href="#hex" x="35" y="20"/>'
            '<use href="#hex" x="60" y="20"/>'
        )

        self.assertEqual(slide_xml.count('name="Polygon '), 3)
        self.assertIn(f'<a:off x="{10 * EMU_PER_PX}" y="{20 * EMU_PER_PX}"/>', slide_xml)
        self.assertIn(f'<a:off x="{35 * EMU_PER_PX}" y="{20 * EMU_PER_PX}"/>', slide_xml)
        self.assertIn(f'<a:off x="{60 * EMU_PER_PX}" y="{20 * EMU_PER_PX}"/>', slide_xml)

    def test_style_rgb_fallback_over_var_fill_and_stroke(self) -> None:
        slide_xml = _render_svg(
            '<rect x="10" y="20" width="30" height="40" '
            'fill="var(--bg)" stroke="var(--border)" '
            'style="fill:rgb(230, 241, 251);stroke:rgba(24, 95, 165, 0.5);stroke-width:0.5px"/>'
        )

        self.assertIn('prst="rect"', slide_xml)
        self.assertIn('val="E6F1FB"', slide_xml)
        self.assertIn('<a:ln w="4762"', slide_xml)
        self.assertIn('val="185FA5"><a:alpha val="50000"/>', slide_xml)

    def test_text_style_rgb_and_px_font_size_are_preserved(self) -> None:
        slide_xml = _render_svg(
            '<text x="100" y="50" '
            'style="fill:rgb(12, 68, 124);font-size:14px;font-family:Arial">Hello</text>'
        )

        self.assertIn('sz="1050"', slide_xml)
        self.assertIn('val="0C447C"', slide_xml)

    def test_marker_end_becomes_arrowhead(self) -> None:
        slide_xml = _render_svg(
            '<defs><marker id="arrow" viewBox="0 0 10 10"/></defs>'
            '<line x1="10" y1="20" x2="50" y2="20" '
            'stroke="#000000" stroke-width="1.2" marker-end="url(#arrow)"/>'
        )

        self.assertIn('<a:tailEnd type="arrow" w="med" len="med"/>', slide_xml)

    def test_rect_with_rx_becomes_round_rect(self) -> None:
        slide_xml = _render_svg(
            '<rect x="10" y="20" width="30" height="40" rx="6" fill="#ffffff"/>'
        )

        self.assertIn('prst="roundRect"', slide_xml)


if __name__ == '__main__':
    unittest.main()
