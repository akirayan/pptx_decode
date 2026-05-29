#!/usr/bin/env python3

"""
make_template.py
Generate a minimal demo template PPTX — no copyright issues.

The generated file (demo_template.pptx) is widescreen (33.867 x 19.05 cm, 16:9)
and has python-pptx's standard 11 layouts.  Use it as a starting point when you
do not have a company template available.

Run:
  python3 make_template.py
  python3 make_template.py my_template.pptx

After generating, check the layout indexes:
  pptx_decode.py demo_template.pptx layout
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from pptx import Presentation
from pptx.util import Cm

OUTPUT = Path(__file__).parent / "demo_template.pptx"

# Widescreen 16:9  (same as most modern PowerPoint templates)
SLIDE_W = Cm(33.867)
SLIDE_H = Cm(19.05)


def build_template(output=None):
    output = output or OUTPUT

    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    prs.save(str(output))

    print(f"Generated: {output}")
    print(f"  Slide size: {prs.slide_width.cm:.3f} x {prs.slide_height.cm:.3f} cm  (16:9)")
    print()
    print("Layout indexes:")
    for i, layout in enumerate(prs.slide_layouts):
        phs = [str(ph.placeholder_format.type).split('(')[0].strip()
               for ph in layout.placeholders]
        print(f"  {i}: {layout.name:<28}  {phs}")
    print()
    print(f"Set in your build script:")
    print(f"  TEMPLATE = '{output.name}'")
    print(f"  LAYOUT_TITLE   = 0   # Title Slide")
    print(f"  LAYOUT_CONTENT = 1   # Title and Content")
    print(f"  LAYOUT_BLANK   = 6   # Blank")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else None
    build_template(Path(out) if out else None)
