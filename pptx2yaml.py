#!/usr/bin/env python3

"""
pptx2yaml.py
Convert PPTX to YAML (for yaml2pptx.py round-trip)

Usage:
  pptx2yaml.py file.pptx
  pptx2yaml.py file.pptx -o output.yaml
"""

import sys
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import yaml
from pptx import Presentation
from pptx_yaml import decode_to_dict, print_yaml


def main():
    parser = argparse.ArgumentParser(
        description='Convert PPTX to YAML',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  pptx2yaml.py file.pptx                   # stdout
  pptx2yaml.py file.pptx -o slides.yaml    # file
  pptx2yaml.py file.pptx | yaml2pptx.py - out.pptx --template tmpl.pptx
        """,
    )
    parser.add_argument('pptx', help='Input PPTX file')
    parser.add_argument('-o', '--output', default=None,
                        help='Output YAML file (default: stdout)')
    args = parser.parse_args()

    prs = Presentation(args.pptx)

    if args.output:
        data = decode_to_dict(prs)
        with open(args.output, 'w', encoding='utf-8') as f:
            yaml.dump(data, f, allow_unicode=True,
                      default_flow_style=False, sort_keys=False)
        print(f"Saved: {args.output}", file=sys.stderr)
    else:
        print_yaml(prs)


if __name__ == '__main__':
    main()
