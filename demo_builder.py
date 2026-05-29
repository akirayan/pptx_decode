#!/usr/bin/env python3

"""
demo_builder.py
pptx_components.py ライブラリの使用例集

Run:
  python3 demo_builder.py
  python3 demo_builder.py my_output.pptx

Slide list:
  1. タイトル（add_section_slide）
  2. スライド追加の4方法（add_blank_slide / add_title_slide / add_content_slide / add_section_slide）
  3. 基本シェイプ（add_rect / add_text / add_line）
  4. テキスト書式（size / bold / color / align / font_name）
  5. カラー定数 ＆ フォントサイズ定数
  6. add_card
  7. add_table ＆ add_text_box
  8. add_content_header ＆ add_bullet_block
  9. add_note ＆ add_picture
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from pptx_components import *

# ============================================================
# Template
#   None  → python-pptx デフォルト（25.4 x 19.05cm）
#   path  → テンプレートファイルを指定
#
# テンプレートを使う場合は、まず以下でLayout Indexを確認する：
#   pptx_decode.py template.pptx layout
# ============================================================
TEMPLATE = None   # 例: "cybersolutions_template.pptx"

# ============================================================
# Layout indexes
#   テンプレートなし（TEMPLATE = None）の場合の標準インデックス：
#     0: Title Slide       CENTER_TITLE + SUBTITLE
#     1: Title and Content TITLE + BODY
#     2: Section Header    TITLE + BODY
#     6: Blank             (no content placeholder)
#
#   テンプレートを使う場合は pptx_decode.py で確認した番号に変更する
# ============================================================
LAYOUT_TITLE   = 0
LAYOUT_CONTENT = 1
LAYOUT_SECTION = 2
LAYOUT_BLANK   = 6

OUTPUT = Path(__file__).parent / "demo_output.pptx"

FONT = "Meiryo UI"   # font used in CS-brand slides


# ============================================================
# Internal helper
# ============================================================

def _section_bar(slide, text):
    """Small label in top-right corner to identify each demo slide."""
    add_text(slide, text, 17.0, 0.2, 8.0, 0.7,
             size=FS_H6, color=GRAY_LINE, align=PP_ALIGN.RIGHT)


# ============================================================
# Slide 1 — Title
# ============================================================

def slide_01_title(prs):
    # add_section_slide handles both TITLE and CENTER_TITLE placeholders
    add_section_slide(
        prs, LAYOUT_TITLE,
        "pptx_components.py",
        "Python PPTX ライブラリ 使用例集",
    )


# ============================================================
# Slide 2 — スライド追加の4方法
# ============================================================

def slide_02_slide_types(prs):
    slide = add_blank_slide(prs, LAYOUT_BLANK)
    _section_bar(slide, "スライド追加")

    add_text(slide, "スライド追加：4つの関数",
             1.0, 0.4, 23.0, 0.9, size=FS_H2, bold=True, color=GRAY_DARK)
    add_line(slide, 1.0, 1.4, 24.4, 1.4, color=GRAY_LINE, width=0.5)

    rows_info = [
        ("add_title_slide",
         "TITLE + SUBTITLE",
         "TITLE プレースホルダーを持つ\nレイアウト専用。",
         "add_title_slide(prs, layout, title, subtitle)"),
        ("add_content_slide",
         "TITLE + BODY",
         "タイトルと本文プレースホルダーを\n持つレイアウト用。",
         "add_content_slide(prs, layout, title, body)"),
        ("add_blank_slide",
         "プレースホルダーなし",
         "全シェイプを削除した空スライド。\nadd_rect / add_text で自由配置。",
         "add_blank_slide(prs, layout)"),
        ("add_section_slide",
         "CENTER_TITLE / TITLE + SUBTITLE",
         "タイトル・セクション区切りスライド。\nCENTER_TITLE を優先して設定する。",
         "add_section_slide(prs, layout, title, subtitle)"),
    ]

    xs = [1.0, 7.1, 13.2, 19.3]
    w  = 5.8

    for i, (fn, ph, desc, sig) in enumerate(rows_info):
        x = xs[i]
        add_rect(slide, x, 1.8, w, 15.5, fill=GRAY_LIGHT,
                 line_color=GRAY_LINE, line_width=0.5)
        add_rect(slide, x, 1.8, w, 1.2, fill=CS_RED)
        add_text(slide, fn, x+0.15, 1.9, w-0.3, 1.0,
                 size=FS_H5, bold=True, color=WHITE)
        add_text(slide, ph,   x+0.2, 3.3, w-0.4, 0.9,
                 size=FS_H6, bold=True, color=CS_RED)
        add_text(slide, desc, x+0.2, 4.4, w-0.4, 3.5,
                 size=FS_H6, color=GRAY_DARK)
        add_rect(slide, x+0.2, 8.2, w-0.4, 1.6, fill=GRAY_DARK)
        add_text(slide, sig,  x+0.3, 8.35, w-0.5, 1.3,
                 size=FS_H6, color=GREEN_LIGHT)

    add_note(slide,
             "テンプレートで作ったレイアウトのプレースホルダーを使うか、"
             "add_blank_slide で全部手描きするかを用途で使い分ける")


# ============================================================
# Slide 3 — add_rect / add_text / add_line
# ============================================================

def slide_03_basic_shapes(prs):
    slide = add_blank_slide(prs, LAYOUT_BLANK)
    _section_bar(slide, "基本シェイプ")

    add_text(slide, "add_rect  /  add_text  /  add_line",
             1.0, 0.4, 23.0, 0.9, size=FS_H2, bold=True, color=GRAY_DARK)
    add_line(slide, 1.0, 1.4, 24.4, 1.4, color=GRAY_LINE, width=0.5)

    # --- add_rect ---
    add_text(slide, "add_rect", 1.0, 1.7, 12.0, 0.7,
             size=FS_H4, bold=True, color=CS_RED)

    samples_rect = [
        (1.0,  2.5, 5.5, 2.5, CS_RED_PALE,  None,     0,   "fill=CS_RED_PALE"),
        (7.0,  2.5, 5.5, 2.5, BLUE_LIGHT,   BLUE,     1.5, "fill + line_color + line_width"),
        (13.0, 2.5, 5.5, 2.5, None,         GRAY_MID, 1.0, "fill=None（透明）"),
        (1.0,  5.5, 5.5, 2.5, GREEN_LIGHT,  GREEN,    0.5, "fill + line_color=GREEN"),
        (7.0,  5.5, 5.5, 2.5, ORANGE_LIGHT, ORANGE,   2.0, "line_width=2.0"),
        (13.0, 5.5, 5.5, 2.5, PURPLE_LIGHT, PURPLE,   1.0, "fill=PURPLE_LIGHT"),
    ]
    for x, y, w, h, fill, lc, lw, label in samples_rect:
        add_rect(slide, x, y, w, h, fill=fill,
                 line_color=lc, line_width=lw)
        add_text(slide, label, x+0.15, y+h-0.8, w-0.3, 0.7,
                 size=FS_H6, color=GRAY_MID)

    # --- add_line ---
    add_text(slide, "add_line", 1.0, 8.8, 12.0, 0.7,
             size=FS_H4, bold=True, color=BLUE)

    lines = [
        (1.0, 9.8, 18.5, 9.8,  GRAY_LINE, 0.5, "GRAY_LINE / width=0.5"),
        (1.0, 11.0, 18.5, 11.0, CS_RED,   2.0, "CS_RED / width=2.0"),
        (1.0, 12.2, 18.5, 12.2, BLUE,     1.0, "BLUE / width=1.0"),
        (1.0, 13.4, 18.5, 13.4, GREEN,    1.5, "GREEN / width=1.5"),
    ]
    for x1, y1, x2, y2, color, width, label in lines:
        add_line(slide, x1, y1, x2, y2, color=color, width=width)
        add_text(slide, label, 18.8, y1-0.15, 5.5, 0.7,
                 size=FS_H6, color=GRAY_MID)

    add_note(slide,
             "add_rect(slide, x, y, w, h, fill, line_color=None, line_width=0.5)  |  "
             "add_line(slide, x1, y1, x2, y2, color=GRAY_LINE, width=1.0)")


# ============================================================
# Slide 4 — add_text テキスト書式
# ============================================================

def slide_04_text_styles(prs):
    slide = add_blank_slide(prs, LAYOUT_BLANK)
    _section_bar(slide, "テキスト書式")

    add_text(slide, "add_text — サイズ・太字・色・揃え・フォント",
             1.0, 0.4, 23.0, 0.9, size=FS_H2, bold=True, color=GRAY_DARK)
    add_line(slide, 1.0, 1.4, 24.4, 1.4, color=GRAY_LINE, width=0.5)

    samples = [
        # (y,    text,                              size,  bold,  color,      align,             font_name)
        (1.7,  "size=FS_H0  (32pt)",                FS_H0, False, BLACK,      PP_ALIGN.LEFT,     None),
        (3.1,  "size=FS_H1  bold=True  color=CS_RED", FS_H1, True, CS_RED,   PP_ALIGN.LEFT,     None),
        (4.4,  "size=FS_H2  color=BLUE",              FS_H2, False, BLUE,     PP_ALIGN.LEFT,     None),
        (5.5,  "size=FS_H3  color=GREEN",             FS_H3, False, GREEN,    PP_ALIGN.LEFT,     None),
        (6.4,  "size=FS_H4  color=ORANGE",            FS_H4, False, ORANGE,   PP_ALIGN.LEFT,     None),
        (7.2,  "size=FS_H5  color=PURPLE",            FS_H5, False, PURPLE,   PP_ALIGN.LEFT,     None),
        (7.9,  "size=FS_H6  color=GRAY_MID",          FS_H6, False, GRAY_MID, PP_ALIGN.LEFT,     None),
        (9.0,  "align=PP_ALIGN.CENTER",               FS_H3, False, GRAY_DARK,PP_ALIGN.CENTER,   None),
        (10.0, "align=PP_ALIGN.RIGHT",                FS_H3, False, GRAY_DARK,PP_ALIGN.RIGHT,    None),
        (11.2, "font_name='Meiryo UI'  size=FS_H3",   FS_H3, False, GRAY_DARK,PP_ALIGN.LEFT,     FONT),
        (12.3, "bold=True  font_name='Meiryo UI'",    FS_H3, True,  CS_RED,   PP_ALIGN.LEFT,     FONT),
    ]

    for y, text, size, bold, color, align, font_name in samples:
        add_text(slide, text, 1.0, y, 22.5, 1.1,
                 size=size, bold=bold, color=color,
                 align=align, font_name=font_name)

    add_note(slide,
             "add_text(slide, text, x, y, w, h, size=12, bold=False, "
             "color=BLACK, align=PP_ALIGN.LEFT, font_name=None)")


# ============================================================
# Slide 5 — カラー定数 ＆ フォントサイズ定数
# ============================================================

def slide_05_colors(prs):
    slide = add_blank_slide(prs, LAYOUT_BLANK)
    _section_bar(slide, "定数一覧")

    add_text(slide, "カラー定数  /  フォントサイズ定数",
             1.0, 0.4, 23.0, 0.9, size=FS_H2, bold=True, color=GRAY_DARK)
    add_line(slide, 1.0, 1.4, 24.4, 1.4, color=GRAY_LINE, width=0.5)

    swatches = [
        ("CS_RED_DARK",  CS_RED_DARK,  WHITE),
        ("CS_RED",       CS_RED,       WHITE),
        ("CS_RED_MID",   CS_RED_MID,   WHITE),
        ("CS_RED_LIGHT", CS_RED_LIGHT, GRAY_DARK),
        ("CS_RED_PALE",  CS_RED_PALE,  GRAY_DARK),
        ("BLUE",         BLUE,         WHITE),
        ("BLUE_LIGHT",   BLUE_LIGHT,   GRAY_DARK),
        ("GREEN",        GREEN,        WHITE),
        ("GREEN_LIGHT",  GREEN_LIGHT,  GRAY_DARK),
        ("ORANGE",       ORANGE,       WHITE),
        ("ORANGE_LIGHT", ORANGE_LIGHT, GRAY_DARK),
        ("PURPLE",       PURPLE,       WHITE),
        ("PURPLE_LIGHT", PURPLE_LIGHT, GRAY_DARK),
        ("GRAY_DARK",    GRAY_DARK,    WHITE),
        ("GRAY_MID",     GRAY_MID,     WHITE),
        ("GRAY_LIGHT",   GRAY_LIGHT,   GRAY_DARK),
        ("GRAY_LINE",    GRAY_LINE,    GRAY_DARK),
        ("BLACK",        BLACK,        WHITE),
        ("WHITE",        WHITE,        GRAY_DARK),
    ]

    cols = 4
    sw, sh = 5.5, 1.3
    gx, gy = 0.3, 0.15
    x0, y0 = 1.0, 1.8

    for i, (name, fill, tc) in enumerate(swatches):
        col = i % cols
        row = i // cols
        x = x0 + col * (sw + gx)
        y = y0 + row * (sh + gy)
        add_rect(slide, x, y, sw, sh, fill=fill,
                 line_color=GRAY_LINE, line_width=0.3)
        add_text(slide, name, x+0.15, y+0.28, sw-0.3, 0.75,
                 size=FS_H5, color=tc)

    # Font size constants
    add_line(slide, 1.0, 15.6, 24.4, 15.6, color=GRAY_LINE, width=0.5)
    add_text(slide, "フォントサイズ定数",
             1.0, 15.8, 12.0, 0.7, size=FS_H4, bold=True, color=GRAY_DARK)

    fs_items = [
        ("FS_H0=32", FS_H0), ("FS_H1=24", FS_H1), ("FS_H2=20", FS_H2),
        ("FS_H3=16", FS_H3), ("FS_H4=14", FS_H4), ("FS_H5=12", FS_H5),
        ("FS_H6=10", FS_H6),
    ]
    for i, (label, size) in enumerate(fs_items):
        add_text(slide, label, 1.0 + i * 3.3, 16.7, 3.1, 0.9,
                 size=size, color=GRAY_DARK)


# ============================================================
# Slide 6 — add_card
# ============================================================

def slide_06_card(prs):
    slide = add_blank_slide(prs, LAYOUT_BLANK)
    _section_bar(slide, "add_card")

    add_text(slide, "add_card — タイトル ＋ 箇条書きカード",
             1.0, 0.4, 23.0, 0.9, size=FS_H2, bold=True, color=GRAY_DARK)
    add_line(slide, 1.0, 1.4, 24.4, 1.4, color=GRAY_LINE, width=0.5)

    cards = [
        ("CS_RED_PALE / CS_RED",   CS_RED_PALE,  CS_RED,
         ["項目 A：説明テキスト", "項目 B：説明テキスト", "項目 C：説明テキスト"]),
        ("BLUE_LIGHT / BLUE",      BLUE_LIGHT,   BLUE,
         ["特徴 1：説明テキスト", "特徴 2：説明テキスト", "特徴 3：説明テキスト"]),
        ("GREEN_LIGHT / GREEN",    GREEN_LIGHT,  GREEN,
         ["ポイント X", "ポイント Y", "ポイント Z"]),
        ("ORANGE_LIGHT / ORANGE",  ORANGE_LIGHT, ORANGE,
         ["注意事項 α", "注意事項 β"]),
    ]

    xs = [1.0, 7.1, 13.2, 19.3]
    for i, (label, fill, tc, items) in enumerate(cards):
        add_card(slide, label, items, xs[i], 2.0, 5.8, 14.0,
                 fill=fill, title_color=tc)

    add_text(slide,
             "add_card(slide, title, items, x, y, w, h, fill=GRAY_LIGHT, title_color=BLACK)",
             1.0, 16.8, 23.4, 0.8, size=FS_H6, color=GRAY_MID, align=PP_ALIGN.CENTER)


# ============================================================
# Slide 7 — add_table ＆ add_text_box
# ============================================================

def slide_07_table_textbox(prs):
    slide = add_blank_slide(prs, LAYOUT_BLANK)
    _section_bar(slide, "add_table / add_text_box")

    add_text(slide, "add_table  /  add_text_box",
             1.0, 0.4, 23.0, 0.9, size=FS_H2, bold=True, color=GRAY_DARK)
    add_line(slide, 1.0, 1.4, 24.4, 1.4, color=GRAY_LINE, width=0.5)

    # add_table
    add_text(slide, "add_table", 1.0, 1.7, 12.0, 0.8,
             size=FS_H3, bold=True, color=CS_RED)

    rows = [
        ["関数",               "主なパラメータ",                 "説明"],
        ["add_text",          "x,y,w,h / size / bold / color", "テキストボックス"],
        ["add_rect",          "x,y,w,h / fill / line_color",   "矩形（塗り＋枠線）"],
        ["add_line",          "x1,y1,x2,y2 / color / width",   "直線"],
        ["add_card",          "title / items / fill",           "カードコンポーネント"],
        ["add_table",         "rows / header_fill / alt_fill",  "テーブル"],
        ["add_content_header","title / subtitle / font_name",   "CSヘッダー"],
        ["add_bullet_block",  "title / bullets / bg",           "箇条書きブロック"],
        ["add_note",          "text / y=17.75",                 "フッターノート"],
    ]
    add_table(slide, rows, 1.0, 2.6, 23.4, 10.0,
              header_fill=CS_RED, header_text_color=WHITE,
              alt_fill=CS_RED_PALE, font_size=FS_H6)

    # add_text_box
    add_text(slide, "add_text_box", 1.0, 13.3, 12.0, 0.8,
             size=FS_H3, bold=True, color=BLUE)

    add_text_box(slide, "bold=True  size=16  color=CS_RED",
                 1.0, 14.2, 22.4, 1.0,
                 font_size=16, bold=True, color=CS_RED)
    add_text_box(slide, "italic=True  size=14  color=BLUE  — add_text_box は italic をサポート",
                 1.0, 15.4, 22.4, 0.9,
                 font_size=14, italic=True, color=BLUE)
    add_text_box(slide, "align=CENTER  color=GRAY_MID",
                 1.0, 16.5, 23.4, 0.9,
                 font_size=14, color=GRAY_MID, align=PP_ALIGN.CENTER)


# ============================================================
# Slide 8 — add_content_header ＆ add_bullet_block
# ============================================================

def slide_08_cs_components(prs):
    slide = add_blank_slide(prs, LAYOUT_BLANK)

    # add_content_header: fixed-position CS-brand header
    add_content_header(slide,
                       "add_content_header  /  add_bullet_block",
                       "CS ブランドスタイルのヘッダーと箇条書きブロック",
                       font_name=FONT)

    blocks = [
        ("✅ 良い点",
         ["使いやすい関数 API", "再利用可能なコンポーネント", "YAML ラウンドトリップ対応"],
         GREEN,  GREEN_LIGHT),
        ("⚠️ 懸念点",
         ["座標は手動指定", "テンプレート依存あり"],
         CS_RED, CS_RED_LIGHT),
        ("🔧 改善案",
         ["グリッドレイアウト helper", "型ヒント追加", "テスト拡充"],
         BLUE,   BLUE_LIGHT),
        ("📌 アクション",
         ["pptx_build.py で自動化", "CSV/JSON 入力対応"],
         ORANGE, ORANGE_LIGHT),
    ]

    xs = [1.0, 7.1, 13.2, 19.3]
    for i, (title, bullets, tc, bg) in enumerate(blocks):
        add_bullet_block(slide, title, bullets,
                         xs[i], 3.0, 5.8, 14.3,
                         title_color=tc, bg=bg, font_name=FONT)

    add_note(slide,
             "add_content_header(slide, title, subtitle, font_name)  |  "
             "add_bullet_block(slide, title, bullets, l, t, w, h, "
             "title_color, bg, bullet_color, font_name)")


# ============================================================
# Slide 9 — add_note ＆ add_picture
# ============================================================

def slide_09_note_picture(prs):
    slide = add_blank_slide(prs, LAYOUT_BLANK)
    _section_bar(slide, "add_note / add_picture")

    add_text(slide, "add_note  /  add_picture",
             1.0, 0.4, 23.0, 0.9, size=FS_H2, bold=True, color=GRAY_DARK)
    add_line(slide, 1.0, 1.4, 24.4, 1.4, color=GRAY_LINE, width=0.5)

    # add_note
    add_text(slide, "add_note", 1.0, 1.7, 12.0, 0.8,
             size=FS_H3, bold=True, color=CS_RED)
    add_text(slide,
             "add_note() はスライド下部（y=17.75）に小さなフッターテキストを追加する。\n"
             "中央揃え、size=8.5pt、color=GRAY_MID で描画される。",
             1.0, 2.6, 22.4, 2.0, size=FS_H4, color=GRAY_DARK)
    add_rect(slide, 1.0, 4.8, 22.4, 1.4, fill=GRAY_LIGHT,
             line_color=GRAY_LINE, line_width=0.5)
    add_text(slide, "add_note(slide, text, y=17.75)", 1.2, 5.0, 22.0, 0.7,
             size=FS_H5, color=CS_RED)
    add_text(slide, "→ ここに表示されるのが add_note のサンプルテキスト（下部を確認）",
             1.2, 5.7, 22.0, 0.7, size=FS_H5, color=GRAY_MID)

    # add_picture
    add_text(slide, "add_picture", 1.0, 7.2, 12.0, 0.8,
             size=FS_H3, bold=True, color=BLUE)
    add_text(slide,
             "add_picture(slide, image_path, x, y, w=None, h=None)\n\n"
             "・w / h 両方省略 → 原寸で挿入\n"
             "・w のみ指定    → 幅に合わせてアスペクト比を維持\n"
             "・h のみ指定    → 高さに合わせてアスペクト比を維持\n"
             "・w / h 両方指定 → 指定サイズに引き伸ばし",
             1.0, 8.2, 22.4, 5.0, size=FS_H4, color=GRAY_DARK)

    # Try to show a real picture if a sample image exists nearby
    sample_img = Path(__file__).parent.parent / "cert_platform" / "cybersolutions_logo.jpg"
    if sample_img.exists():
        add_text(slide, "logo サンプル（w=6.0cm）", 1.0, 13.5, 12.0, 0.7,
                 size=FS_H5, color=GRAY_MID)
        add_picture(slide, sample_img, 1.0, 14.3, w=6.0)
    else:
        add_rect(slide, 1.0, 13.5, 8.0, 3.5, fill=GRAY_LIGHT,
                 line_color=GRAY_LINE, line_width=1.0)
        add_text(slide, "（画像ファイルなし）", 1.2, 14.7, 7.5, 1.0,
                 size=FS_H5, color=GRAY_MID, align=PP_ALIGN.CENTER)

    add_note(slide, "※ これが add_note() のサンプル —— スライド下部に自動配置される（y=17.75、中央揃え、8.5pt）")


# ============================================================
# Main
# ============================================================

def build(output=None):
    output = output or OUTPUT
    prs = create_ppt(TEMPLATE)

    slide_01_title(prs)
    slide_02_slide_types(prs)
    slide_03_basic_shapes(prs)
    slide_04_text_styles(prs)
    slide_05_colors(prs)
    slide_06_card(prs)
    slide_07_table_textbox(prs)
    slide_08_cs_components(prs)
    slide_09_note_picture(prs)

    save_presentation(prs, str(output))
    print(f"Generated: {output}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else None
    build(out)
