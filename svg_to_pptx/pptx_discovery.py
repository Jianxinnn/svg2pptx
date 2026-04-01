"""Find SVG and notes files from a project directory, a plain folder, or a single SVG."""

from __future__ import annotations

import re
from pathlib import Path


SVG_EXTENSIONS = ('.svg',)


def find_svg_files(
    input_path: Path,
    source: str = 'output',
) -> tuple[list[Path], str]:
    """Find SVG files from a project directory, folder, or single SVG file.

    Args:
        input_path: Project directory path, generic folder path, or single SVG file.
        source: SVG source directory alias or name when input_path is a project.
            - 'output': svg_output (original version)
            - 'final': svg_final (post-processed, recommended)
            - or any subdirectory name

    Returns:
        (list_of_svg_files, actual_directory_name) tuple.
    """
    if input_path.is_file():
        if input_path.suffix.lower() in SVG_EXTENSIONS:
            return [input_path], input_path.parent.name
        return [], ''

    if not input_path.exists() or not input_path.is_dir():
        return [], ''

    dir_map = {
        'output': 'svg_output',
        'final': 'svg_final',
    }

    dir_name = dir_map.get(source, source)
    candidate_dir = input_path / dir_name

    if candidate_dir.exists() and candidate_dir.is_dir():
        return sorted(candidate_dir.glob('*.svg')), dir_name

    direct_svgs = sorted(input_path.glob('*.svg'))
    if direct_svgs:
        return direct_svgs, input_path.name

    if dir_name != 'svg_output':
        fallback_dir = input_path / 'svg_output'
        if fallback_dir.exists() and fallback_dir.is_dir():
            print("  Warning: requested SVG directory does not exist, using svg_output")
            return sorted(fallback_dir.glob('*.svg')), 'svg_output'

    return [], ''


def find_notes_files(
    project_path: Path,
    svg_files: list[Path] | None = None,
) -> dict[str, str]:
    """Find notes files and map them to SVG files.

    Supports two matching modes (mixed matching supported):
    1. Match by filename (priority): notes/01_cover.md -> 01_cover.svg
    2. Match by index (backward compatible): notes/slide01.md -> 1st SVG

    Args:
        project_path: Project directory path.
        svg_files: SVG file list (for filename matching).

    Returns:
        Dict mapping SVG filename stem to notes content.
    """
    notes_dir = project_path / 'notes'
    notes: dict[str, str] = {}

    if not notes_dir.exists():
        return notes

    svg_stems_mapping: dict[str, int] = {}
    svg_index_mapping: dict[int, str] = {}
    if svg_files:
        for i, svg_path in enumerate(svg_files, 1):
            svg_stems_mapping[svg_path.stem] = i
            svg_index_mapping[i] = svg_path.stem

    for notes_file in notes_dir.glob('*.md'):
        try:
            with open(notes_file, 'r', encoding='utf-8') as f:
                content = f.read().strip()
            if not content:
                continue

            stem = notes_file.stem

            # Try index-based matching (backward compat with slide01.md format)
            match = re.search(r'slide[_]?(\d+)', stem)
            if match:
                index = int(match.group(1))
                mapped_stem = svg_index_mapping.get(index)
                if mapped_stem:
                    notes[mapped_stem] = content

            # Filename-based matching (overrides index-based)
            if stem in svg_stems_mapping:
                notes[stem] = content
        except Exception:
            pass

    return notes
