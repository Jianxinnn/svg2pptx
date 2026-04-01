#!/usr/bin/env python3
"""svg2ppt standalone CLI wrapper.

Delegates to the svg_to_pptx package:
    python3 svg2ppt/svg_to_pptx.py <input_path>
"""

import sys
from pathlib import Path

# Ensure the scripts directory is on sys.path so the package can be found
sys.path.insert(0, str(Path(__file__).resolve().parent))

from svg_to_pptx import main

if __name__ == '__main__':
    main()
