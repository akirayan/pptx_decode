# PPTX ライブラリ 関数一覧

<idx>pptx</idx>

## ファイル構成

| ファイル | 役割 |
|---|---|
| `pptx_core.py` | 低レベルの基本操作（Presentation / Slide / Shape / Text） |
| `pptx_components.py` | 再利用可能な高レベルコンポーネント（カード・テーブルなど） |
| `pptx_decode.py` | CLIツール：既存PPTXの構造を人間が読める形式で表示 |
| `pptx_yaml.py` | PPTXスライドをYAML構造にデコードするライブラリ |
| `pptx2yaml.py` | CLIツール：PPTX → YAML 変換 |
| `yaml2pptx.py` | CLIツール：YAML → PPTX 変換 |

---

## pptx_core.py — コア操作

### Presentation（プレゼンテーション）

| 関数 | シグネチャ | 説明 |
|---|---|---|
| `create_ppt` | `(template=None)` | テンプレートからPresentationを生成（スライドは削除、レイアウト/マスターを保持） |
| `save_ppt` | `(prs, path)` | PPTXファイルを保存 |
| `create_presentation` | `(template=None)` | `Presentation`オブジェクトを生成（ラップなし） |
| `save_presentation` | `(prs, path)` | PPTXファイルを保存 |
| `remove_all_slides` | `(prs)` | 全スライドを削除（テーマ・レイアウトは保持） |

### Layout / Slide（レイアウト・スライド）

| 関数 | シグネチャ | 説明 |
|---|---|---|
| `list_layouts` | `(prs)` | `[(index, name)]` 形式でレイアウト一覧を返す |
| `get_layout_by_name` | `(prs, name)` | 名前でレイアウトを取得（見つからなければ例外） |
| `add_slide` | `(prs, layout_index)` | レイアウトインデックスを指定してスライドを追加 |

### Shape / Text（テキストボックス・テキスト）

| 関数 | シグネチャ | 説明 |
|---|---|---|
| `clear_text` | `(shape)` | shapeのテキストをクリアして`text_frame`を返す |
| `remove_all_items` | `(slide)` | スライド上の全シェイプ（プレースホルダー含む）を削除 |
| `add_paragraph` | `(tf, level=0)` | `text_frame`に段落を追加 |
| `add_run` | `(p, text)` | 段落にテキストrunを追加 |
| `apply_paragraphs` | `(shape, paragraphs)` | 構造化データ（辞書リスト）から段落・runをまとめて適用 |

`apply_paragraphs` のデータ形式:
```python
paragraphs = [
    {"level": 0, "runs": [{"text": "Hello"}, {"text": "World"}]}
]
```

### Placeholder（プレースホルダー）

| 関数 | シグネチャ | 説明 |
|---|---|---|
| `set_placeholder` | `(slide, ptype, text, idx=None)` | タイプ（`PH.TITLE` / `PH.BODY` など）でプレースホルダーにテキストをセット |

### テーマ・デコード

| 関数 | シグネチャ | 説明 |
|---|---|---|
| `get_theme_info` | `(prs)` | テーマのフォント情報を辞書で返す（例: `{"major_font": "Calibri"}`） |
| `get_effective_font` | `(font, theme)` | テーマを考慮した実効フォント情報を辞書で返す |
| `emu_to_cm` | `(v)` | EMU単位をcmに変換 |

### デバッグ・構造表示

| 関数 | シグネチャ | 説明 |
|---|---|---|
| `list_shape_details` | `(shape, theme=None, indent="  ")` | shapeの座標・テキスト・フォント情報を出力 |
| `list_layout_details` | `(prs)` | 全レイアウトの形状を出力 |
| `list_slide_details` | `(prs)` | 全スライドの形状を出力 |
| `list_master_details` | `(prs)` | マスターの形状を出力 |
| `list_theme_details` | `(prs)` | テーマ（フォント・カラー）を出力 |
| `list_full_structure` | `(prs)` | テーマ→マスター→レイアウト→スライドを全出力 |

---

## pptx_components.py — 高レベルコンポーネント

### カラー定数

| 定数 | 色 |
|---|---|
| `CS_RED_DARK` / `CS_RED` / `CS_RED_MID` / `CS_RED_LIGHT` / `CS_RED_PALE` | 赤系 |
| `WHITE` / `BLACK` / `GRAY_DARK` / `GRAY_MID` / `GRAY_LIGHT` / `GRAY_LINE` | モノクロ系 |
| `BLUE` / `BLUE_LIGHT` | 青系 |
| `GREEN` / `GREEN_LIGHT` | 緑系 |
| `ORANGE` / `ORANGE_LIGHT` | オレンジ系 |
| `PURPLE` / `PURPLE_LIGHT` | 紫系 |

### フォントサイズ定数

| 定数 | サイズ |
|---|---|
| `FS_H0` | 32pt |
| `FS_H1` | 24pt |
| `FS_H2` | 20pt |
| `FS_H3` | 16pt |
| `FS_H4` | 12pt |
| `FS_H5` | 10pt |

### スライド追加

| 関数 | シグネチャ | 説明 |
|---|---|---|
| `create_slide` | `(prs, layout_index)` | レイアウトを指定してスライドを追加 |
| `add_title_slide` | `(prs, layout_index, title, subtitle=None)` | タイトル・サブタイトルを持つ表紙スライドを追加 |
| `add_content_slide` | `(prs, layout_index, title, body=None)` | タイトルとボディを持つコンテンツスライドを追加 |
| `add_section_slide` | `(prs, layout_index, title, subtitle=None)` | セクション区切りスライドを追加（`CENTER_TITLE`にも対応） |
| `add_blank_slide` | `(prs, layout_index)` | レイアウトのシェイプをすべて除去した空スライドを追加 |

### 基本シェイプ

| 関数 | シグネチャ | 説明 |
|---|---|---|
| `add_rect` | `(slide, x, y, w, h, fill, line_color=None, line_width=0.5)` | 矩形を追加。`fill=None` で透明背景、`line_color` 指定で枠線あり |
| `add_text` | `(slide, text, x, y, w, h, size=12, bold=False, color=BLACK, align=PP_ALIGN.LEFT, font_name=None)` | テキストボックスを追加（折り返しあり）。`font_name` 指定でフォントを明示設定 |
| `add_text_box` | `(slide, text, x, y, w, h, *, font_size=12, bold, italic, color, align)` | 書式オプション付きのテキストボックスを追加 |
| `add_picture` | `(slide, image_path, x, y, w=None, h=None)` | 画像を挿入（`w`/`h` 省略で原寸） |
| `add_line` | `(slide, x1, y1, x2, y2, color=GRAY_LINE, width=1.0)` | 直線を描画 |
| `add_note` | `(slide, text, y=17.75)` | スライド下部にフッターノートを追加（中央揃え・8.5pt） |

### テーブル

| 関数 | シグネチャ | 説明 |
|---|---|---|
| `add_table` | `(slide, rows, x, y, w, h, *, header_fill, header_text_color, body_fill, alt_fill, font_size)` | ヘッダー行・交互色オプション付きのテーブルを追加 |

`rows`の形式:
```python
rows = [
    ["ヘッダー1", "ヘッダー2"],   # 1行目がヘッダー
    ["データ1",   "データ2"],
]
```

### 複合コンポーネント

| 関数 | シグネチャ | 説明 |
|---|---|---|
| `add_card` | `(slide, title, items, x, y, w, h, fill=GRAY_LIGHT, title_color=BLACK)` | タイトル＋箇条書きリストのカードを追加 |
| `add_bullet_panel` | `(slide, add_text, add_rect, title, bullets, x, y, w, h, title_color, bg)` | 背景色付きの箇条書きパネルを追加 |
| `add_block` | `(slide, add_text, title, body, x, y, w, h)` | タイトルとテキスト本文のシンプルなブロックを追加 |
| `add_blocks_grid` | `(slide, blocks, add_text, cols=2, x, y, w, h)` | 複数ブロックをグリッド状に並べる |
| `add_header` | `(slide, add_rect, add_text, title, subtitle=None)` | スライド上部のヘッダーエリアを追加 |
| `add_full_background` | `(slide, add_rect, color)` | スライド全面に背景色を塗る |
| `add_agent_slide` | `(prs, layout_index, add_text, add_rect, title, subtitle, good, risks, solutions, priorities)` | Good/Risks/Solutions/Priorityの4パネル構成スライドを追加 |
| `add_content_header` | `(slide, title, subtitle=None, font_name=None)` | CSブランドヘッダーバーを追加（タイトル＋サブタイトル、固定座標） |
| `add_bullet_block` | `(slide, title, bullets, l, t, w, h, title_color=CS_RED, bg=CS_RED_PALE, bullet_color=GRAY_DARK, font_name=None)` | カラーボックス＋タイトル＋箇条書きブロックを追加 |

---

## pptx_decode.py — CLI解析ツール（人間向け表示）

```bash
python pptx_decode.py file.pptx           # full（全情報）
python pptx_decode.py file.pptx layout    # レイアウト一覧
python pptx_decode.py file.pptx slide     # スライド内容
python pptx_decode.py file.pptx master    # マスター情報
python pptx_decode.py file.pptx theme     # テーマ情報
```

---

## pptx2yaml.py / yaml2pptx.py — 変換ツール

```bash
# PPTX → YAML
pptx2yaml.py file.pptx                   # stdout
pptx2yaml.py file.pptx -o slides.yaml    # ファイル出力

# YAML → PPTX
yaml2pptx.py slides.yaml output.pptx
yaml2pptx.py slides.yaml output.pptx --template template.pptx
cat slides.yaml | yaml2pptx.py - output.pptx

# パイプで直接ラウンドトリップ
pptx2yaml.py source.pptx | yaml2pptx.py - rebuilt.pptx --template template.pptx
```

---

## pptx_yaml.py — YAML デコードライブラリ

`pptx2yaml.py` が内部で使用する。直接利用も可能。

### YAML スキーマ

```yaml
template: ./demo_template.pptx   # 省略可。--template オプションより優先度が低い

slides:
  - index: 0
    layout_index: 3
    shapes:
      - fn: add_text        # add_text() に対応
        x: 1.5
        y: 0.5
        w: 26.0
        h: 1.2
        text: タイトルテキスト
        size: 24.0
        bold: true
        color: '#000000'
        font_name: Meiryo UI
        align: center       # left 以外の場合のみ出力

      - fn: add_rect        # add_rect() に対応
        x: 1.02
        y: 3.94
        w: 7.62
        h: 13.21
        fill: '#EEF4FF'     # null = 透明
        line_color: '#2D6BE4'
        line_width: 1.2

      - fn: add_line        # add_line() に対応
        x1: 1.0
        y1: 5.0
        x2: 32.0
        y2: 5.0
        color: '#DDDDDD'
        width: 1.0

      - fn: add_picture     # add_picture() に対応
        x: 1.0
        y: 1.0
        w: 6.0
        h: 2.0

      - fn: set_placeholder # add_title_slide 等のプレースホルダー
        x: 0.873
        y: 8.129
        w: 15.134
        h: 2.22
        placeholder_type: TITLE
        placeholder_idx: 0
        text: タイトルテキスト
```

### 関数

| 関数 | シグネチャ | 説明 |
|---|---|---|
| `decode_to_dict` | `(prs)` | 全スライドを辞書にデコードして返す |
| `print_yaml` | `(prs)` | YAML を stdout に出力 |

---

## yaml2pptx.py — YAML から PPTX を生成

`pptx2yaml.py` の出力を読み込み、PPTX を再構築する。YAML の `fn` キーを `pptx_components.py` の関数に dispatch する。

### fn マッピング

| fn | 呼び出す関数 | 必須キー |
|---|---|---|
| `add_text` | `add_text()` | x, y, w, h, text |
| `add_rect` | `add_rect()` | x, y, w, h, fill |
| `add_line` | `add_line()` | x1, y1, x2, y2 |
| `add_picture` | `add_picture()` | x, y, w, h, image_path |
| `set_placeholder` | `set_placeholder()` | placeholder_type, text |

---

## 基本的な使い方

```python
from pptx_core import create_ppt, save_presentation
from pptx_components import *

# テンプレートから生成
prs = create_ppt("template.pptx")

# タイトルスライド
add_title_slide(prs, 0, "タイトル", "サブタイトル")

# コンテンツスライド
slide = add_content_slide(prs, 3, "ページタイトル")

# カードを追加
add_card(slide, "見出し", ["項目1", "項目2"], x=1, y=3, w=10, h=6, fill=BLUE_LIGHT)

# テーブルを追加
add_table(slide, [["列A", "列B"], ["値1", "値2"]], x=12, y=3, w=15, h=5,
          header_fill=CS_RED, header_text_color=WHITE, alt_fill=CS_RED_PALE)

# 保存
save_presentation(prs, "output.pptx")
```
