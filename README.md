# svg2ppt

Standalone SVG to PPTX converter extracted from `skills/ppt-master`.

It converts:

- a single `.svg` file into a PPTX
- a folder of `.svg` files into a multi-slide PPTX
- a project-like directory using the extracted standalone discovery logic

Default output is a native editable PPTX built with DrawingML shapes.

## Features

- Single SVG to PPTX conversion
- Folder of SVGs to multi-slide PPTX conversion
- Native editable PowerPoint shapes by default
- Optional legacy SVG reference PPTX output
- Auto-detect canvas size from SVG `viewBox`
- Optional speaker notes loading for directory input
- Optional slide transitions
- Office compatibility mode support
- Targeted handling for common dot-pattern SVG cases

## Directory layout

```text
svg2ppt/
├── README.md
├── svg_to_pptx.py
└── svg_to_pptx/
```

Main entry:

```bash
python3 svg2ppt/svg_to_pptx.py <input>
```

## Requirements

Recommended:

- Python 3.10+
- `python-pptx`

Optional compatibility dependencies:

```bash
pip install svglib reportlab
```

If your local environment does not already have the converter dependencies installed, install at least:

```bash
pip install python-pptx svglib reportlab
```

## Quick start

### Convert a single SVG

```bash
python3 svg2ppt/svg_to_pptx.py ./slide.svg
```

### Convert a folder of SVGs

```bash
python3 svg2ppt/svg_to_pptx.py ./svgs
```

### Specify output path

```bash
python3 svg2ppt/svg_to_pptx.py ./svgs -o ./out/result.pptx
```

### Generate legacy SVG reference PPTX only

```bash
python3 svg2ppt/svg_to_pptx.py ./svgs --only legacy
```

### Quiet mode

```bash
python3 svg2ppt/svg_to_pptx.py ./svgs --quiet
```

## Input behavior

svg2ppt accepts these input forms:

1. A single `.svg` file
2. A directory containing `.svg` files directly
3. A project-like directory that can be resolved by the standalone discovery logic

Typical examples:

```bash
python3 svg2ppt/svg_to_pptx.py /Users/me/work/slide.svg
python3 svg2ppt/svg_to_pptx.py /Users/me/work/svgs
python3 svg2ppt/svg_to_pptx.py /Users/me/work/project_dir
```

## Output behavior

By default, svg2ppt generates only one native PPTX.

### When input is a single SVG

Output goes to the SVG file's parent directory.

Example:

```bash
python3 svg2ppt/svg_to_pptx.py /Users/me/Downloads/demo.svg
```

Will generate something like:

```text
/Users/me/Downloads/demo_20260401_153000.pptx
```

### When input is a folder

Output goes into that folder.

Example:

```bash
python3 svg2ppt/svg_to_pptx.py /Users/me/Downloads/svgs
```

Will generate something like:

```text
/Users/me/Downloads/svgs/svgs_20260401_153000.pptx
```

### When `-o` is provided

That exact output path is used.

Example:

```bash
python3 svg2ppt/svg_to_pptx.py ./svgs -o ./build/final.pptx
```

Output:

```text
./build/final.pptx
```

## CLI options

```text
positional arguments:
  project_path           Project directory, SVG folder, or single SVG path

optional arguments:
  -o, --output           Output file path
  -f, --format           Specify canvas format
  -q, --quiet            Quiet mode
  --no-compat            Disable Office compatibility mode
  --only {native,legacy} Generate only native or only legacy output
  --native               Deprecated alias for native-only output
  -t, --transition       Page transition effect
  --transition-duration  Transition duration in seconds
  --auto-advance         Auto-advance interval in seconds
  --no-notes             Disable speaker notes embedding
```

## Canvas format

If `--format` is not provided, svg2ppt tries to infer the canvas from the input name or SVG metadata.

Supported format keys currently include:

- `ppt169`
- `ppt43`
- `wechat`
- `xiaohongshu`
- `moments`
- `story`
- `banner`
- `a4`

Example:

```bash
python3 svg2ppt/svg_to_pptx.py ./svgs --format ppt169
```

## Speaker notes

When the input is a directory, svg2ppt can also read speaker notes if matching notes files exist.

When the input is a single SVG file, notes are skipped automatically.

If you want to disable notes explicitly:

```bash
python3 svg2ppt/svg_to_pptx.py ./svgs --no-notes
```

## Transition settings

You can optionally set slide transitions.

Example:

```bash
python3 svg2ppt/svg_to_pptx.py ./svgs -t fade --transition-duration 0.4
```

Auto-advance example:

```bash
python3 svg2ppt/svg_to_pptx.py ./svgs -t fade --auto-advance 3
```

## Compatibility mode

Compatibility mode is enabled by default.

This is intended to improve behavior across different Office versions.

Disable it if you want pure SVG-oriented output behavior:

```bash
python3 svg2ppt/svg_to_pptx.py ./svgs --no-compat
```

## Current limitations

This is a focused extraction of the existing SVG-to-PPT pipeline, not a full rewrite.

Known limitations include:

- very complex SVG features may still not map perfectly to editable PowerPoint shapes
- some advanced SVG constructs are approximated during conversion
- support is strongest for the patterns already handled by the original converter

## Troubleshooting

### No SVG files found

Check that:

- the input path exists
- the folder actually contains `.svg` files
- the provided path points to the file or folder you intended

### Output looks different from the original SVG

This converter maps SVG into PowerPoint-native structures when possible, so certain effects may be approximated.

If you need a reference-style output instead of editable shapes, try:

```bash
python3 svg2ppt/svg_to_pptx.py ./svgs --only legacy
```

### Some dots or repeated pattern elements are missing

The extractor already includes targeted fixes for common dot-pattern and round-cap dot cases, but very custom SVG constructs may still require additional compatibility work.

## Scope

This standalone tool keeps the original converter core as much as possible, with minimal adaptations for direct folder and single-file usage.
