"""CLI entry point for svg_to_pptx."""

from __future__ import annotations

import sys
import argparse
from datetime import datetime
from pathlib import Path

from .pptx_dimensions import CANVAS_FORMATS, get_project_info
from .pptx_discovery import find_svg_files, find_notes_files
from .pptx_builder import create_pptx_with_native_svg
from .pptx_slide_xml import TRANSITIONS


def main() -> None:
    """CLI entry point for the SVG to PPTX conversion tool."""
    transition_choices = (
        ['none'] + (list(TRANSITIONS.keys()) if TRANSITIONS
                    else ['fade', 'push', 'wipe', 'split', 'strips', 'cover', 'random'])
    )

    parser = argparse.ArgumentParser(
        description='svg2ppt - Convert a folder of SVGs or a single SVG into PPTX',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f'''
Examples:
    %(prog)s ./svgs                         # convert all SVGs in a folder
    %(prog)s ./slide.svg                   # convert a single SVG
    %(prog)s ./svgs -o out.pptx            # explicit output path
    %(prog)s ./svgs --only legacy          # generate legacy SVG reference PPTX only

Transition effects (-t/--transition):
    {', '.join(transition_choices)}

Compatibility mode (enabled by default):
    - Automatically generates PNG fallback images, SVG embedded as extension
    - Compatible with all Office versions (including Office LTSC 2021)
    - Newer Office still displays SVG, older versions display PNG
    - Requires svglib: pip install svglib reportlab
    - Use --no-compat to disable

Speaker notes:
    - Reads notes/*.md only when the input path is a project/folder
    - Single SVG mode skips notes automatically
''',
    )

    parser.add_argument('project_path', type=str, help='Project directory, SVG folder, or single SVG path')
    parser.add_argument('-o', '--output', type=str, default=None, help='Output file path')
    parser.add_argument('-f', '--format', type=str,
                        choices=list(CANVAS_FORMATS.keys()), default=None,
                        help='Specify canvas format')
    parser.add_argument('-q', '--quiet', action='store_true', help='Quiet mode')

    parser.add_argument('--no-compat', action='store_true',
                        help='Disable Office compatibility mode (pure SVG only, requires Office 2019+)')

    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument('--only', type=str, choices=['native', 'legacy'], default=None,
                            help='Only generate one version: native (editable shapes) or legacy (SVG image)')
    mode_group.add_argument('--native', action='store_true', default=False,
                            help='(Deprecated, now default) Convert SVG to native DrawingML shapes')

    parser.add_argument('-t', '--transition', type=str, choices=transition_choices, default='fade',
                        help='Page transition effect (default: fade, use "none" to disable)')
    parser.add_argument('--transition-duration', type=float, default=0.4,
                        help='Transition duration in seconds (default: 0.4)')
    parser.add_argument('--auto-advance', type=float, default=None,
                        help='Auto-advance interval in seconds (default: manual advance)')

    parser.add_argument('--no-notes', action='store_true',
                        help='Disable speaker notes embedding')

    args = parser.parse_args()

    input_path = Path(args.project_path)
    if not input_path.exists():
        print(f"Error: Path does not exist: {input_path}")
        sys.exit(1)

    try:
        project_info = get_project_info(str(input_path))
        project_name = project_info.get('name', input_path.stem if input_path.is_file() else input_path.name)
        detected_format = project_info.get('format')
    except Exception:
        project_name = input_path.stem if input_path.is_file() else input_path.name
        detected_format = None

    canvas_format = args.format
    if canvas_format is None and detected_format and detected_format != 'unknown':
        canvas_format = detected_format

    svg_files, source_dir_name = find_svg_files(input_path)

    if not svg_files:
        print("Error: No SVG files found")
        sys.exit(1)

    only_mode = args.only or 'native'
    gen_native = only_mode == 'native'
    gen_legacy = only_mode == 'legacy'

    if args.native:
        gen_native = True
        gen_legacy = False

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    default_output_dir = input_path.parent if input_path.is_file() else input_path
    if args.output:
        output_base = Path(args.output)
        native_path = output_base
        stem = output_base.stem
        legacy_path = output_base.parent / f"{stem}_svg{output_base.suffix}"
    else:
        native_path = default_output_dir / f"{project_name}_{timestamp}.pptx"
        legacy_path = default_output_dir / f"{project_name}_{timestamp}_svg.pptx"

    native_path.parent.mkdir(parents=True, exist_ok=True)

    verbose = not args.quiet

    enable_notes = not args.no_notes and input_path.is_dir()
    notes: dict[str, str] = {}
    if enable_notes:
        notes = find_notes_files(input_path, svg_files)

    transition = args.transition if args.transition != 'none' else None

    shared_kwargs = dict(
        svg_files=svg_files,
        canvas_format=canvas_format,
        verbose=verbose,
        transition=transition,
        transition_duration=args.transition_duration,
        auto_advance=args.auto_advance,
        use_compat_mode=not args.no_compat,
        notes=notes,
        enable_notes=enable_notes,
    )

    success = True

    if gen_native:
        if verbose:
            print("svg2ppt - Native PPTX")
            print("=" * 50)
            print(f"  Input path: {input_path}")
            print(f"  SVG source: {source_dir_name}")
            print(f"  Output file: {native_path}")
            print()

        ok = create_pptx_with_native_svg(
            output_path=native_path,
            use_native_shapes=True,
            **shared_kwargs,
        )
        success = success and ok

    if gen_legacy:
        if verbose:
            if gen_native:
                print()
                print("-" * 50)
            print("svg2ppt - SVG Reference PPTX")
            print("=" * 50)
            print(f"  Input path: {input_path}")
            print(f"  SVG source: {source_dir_name}")
            print(f"  Output file: {legacy_path}")
            print()

        ok = create_pptx_with_native_svg(
            output_path=legacy_path,
            use_native_shapes=False,
            **shared_kwargs,
        )
        success = success and ok

    sys.exit(0 if success else 1)
