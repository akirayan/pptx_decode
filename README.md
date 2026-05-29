# pptx_decode

Python library and CLI tools for creating and analyzing PowerPoint (PPTX) files.

Instead of fighting with the GUI, write slide content in Python or YAML and generate PPTX from the command line.

## Files

| File | Role |
|---|---|
| `pptx_core.py` | Low-level PPTX operations |
| `pptx_components.py` | Reusable high-level components (shapes, colors, cards, tables) |
| `pptx_decode.py` | CLI: inspect existing PPTX structure |
| `pptx_yaml.py` | Library: decode PPTX slides to YAML |
| `pptx2yaml.py` | CLI: convert PPTX → YAML |
| `yaml2pptx.py` | CLI: convert YAML → PPTX |
| `make_template.py` | Generate a minimal widescreen template (copyright-free) |
| `demo_builder.py` | Usage examples for all library functions |

## Requirements

```bash
pip install python-pptx pyyaml
```

Copy the `.py` files to your project folder and add to `sys.path` if needed.

## Quick Start

### Inspect an existing PPTX

```bash
./pptx_decode.py file.pptx layout   # list layout indexes
./pptx_decode.py file.pptx slide    # show slide content
./pptx_decode.py file.pptx theme    # show theme/font info
```

### Build a PPTX from a Python script

```python
from pptx_components import *

TEMPLATE = "my_template.pptx"   # or None for default
LAYOUT_BLANK = 6

prs = create_ppt(TEMPLATE)

slide = add_blank_slide(prs, LAYOUT_BLANK)
add_text(slide, "Hello World", x=1, y=1, w=20, h=2,
         size=FS_H1, bold=True, color=CS_RED)
add_rect(slide, x=1, y=4, w=10, h=6, fill=BLUE_LIGHT,
         line_color=BLUE, line_width=1.0)

save_presentation(prs, "output.pptx")
```

See `demo_builder.py` for a full 9-slide example covering all functions.

### Build a PPTX from YAML

```yaml
# slides.yaml
template: ./my_template.pptx

slides:
  - layout_index: 6
    shapes:
      - fn: add_text
        x: 1.5
        y: 1.0
        w: 20.0
        h: 1.5
        text: Hello World
        size: 24.0
        bold: true
        color: '#C00000'
```

```bash
yaml2pptx.py slides.yaml output.pptx
```

### Round-trip: PPTX → YAML → PPTX

```bash
pptx2yaml.py source.pptx -o slides.yaml
yaml2pptx.py slides.yaml rebuilt.pptx --template source.pptx
```

### Generate a demo template

```bash
python3 make_template.py        # creates demo_template.pptx (widescreen 16:9)
./pptx_decode.py demo_template.pptx layout   # confirm layout indexes
```

## Design Philosophy

PowerPoint's GUI is hard to control precisely — font sizes, shape positions, and colors require tedious manual adjustment.
This library treats a slide as a canvas: create a blank slide, remove all placeholders, and draw freely using coordinate-based functions (`add_text`, `add_rect`, `add_line`).
Template backgrounds and logos (stored in the slide master layer) are inherited even on blank slides.

## License

MIT
