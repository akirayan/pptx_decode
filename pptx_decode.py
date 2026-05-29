#!/usr/bin/env python3

# pptx_decode.py

import sys
from pptx import Presentation
from pptx_core import (
    list_layout_details,
    list_slide_details,
    list_master_details,
    list_theme_details,
    list_full_structure
)


def usage():
    print("Usage:")
    print("  python pptx_decode.py file.pptx           # full (default)")
    print("  python pptx_decode.py file.pptx layout    # layout details")
    print("  python pptx_decode.py file.pptx slide     # slide details")
    print("  python pptx_decode.py file.pptx master    # master info")
    print("  python pptx_decode.py file.pptx theme     # theme info")


def main():
    if len(sys.argv) < 2:
        usage()
        sys.exit(1)

    path = sys.argv[1]
    mode = "full"

    if len(sys.argv) >= 3:
        mode = sys.argv[2]

    prs = Presentation(path)

    if mode == "layout":
        list_layout_details(prs)

    elif mode == "slide":
        list_slide_details(prs)

    elif mode == "master":
        list_master_details(prs)

    elif mode == "theme":
        list_theme_details(prs)

    elif mode == "full":
        list_full_structure(prs)

    else:
        print(f"Unknown mode: {mode}")
        usage()


if __name__ == "__main__":
    main()
