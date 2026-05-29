# テンプレートの使い方

PowerPointのテンプレートファイル（`.pptx`）を使うことで、ロゴ・フォント・カラースキーム・スライドマスターなどのデザインを引き継いだPPTXを生成できる。

## テンプレートの準備

### Option A: 自社・独自テンプレートを使う

会社のブランドガイドラインに沿ったテンプレートを使う場合は、そのファイルをそのまま指定する。

```
TEMPLATE = "your_company_template.pptx"
```

### Option B: make_template.py でデモ用テンプレートを生成する

テンプレートがない場合や動作確認をしたい場合は、`make_template.py` でワイドスクリーン（33.867 × 19.05 cm、16:9）のシンプルなテンプレートを生成できる。python-pptx だけで生成するため著作権の問題がない。

```bash
python3 make_template.py
```

生成されるファイル: `demo_template.pptx`

> **NOTE:** `demo_template.pptx` は `.gitignore` に含まれる（`*.pptx` を無視）。リポジトリには含めず、使う側で生成する。

## テンプレートとLayout Indexの関係

テンプレートには複数の**レイアウト**が定義されており、スライドを追加するときにどのレイアウトを使うかをインデックス番号で指定する。レイアウトの並び順はテンプレートごとに異なるため、**必ず事前に確認する必要がある**。

この確認手順は自動化できない。

## Step 1: Layout Indexの確認

```bash
pptx_decode.py demo_template.pptx layout
```

出力例（`make_template.py` で生成した `demo_template.pptx` の場合）：

```
=== LAYOUT DETAILS ===
[Layout 0] Title Slide
[Layout 1] Title and Content
[Layout 2] Section Header
[Layout 3] Two Content
[Layout 4] Comparison
[Layout 5] Title Only
[Layout 6] Blank
[Layout 7] Content with Caption
[Layout 8] Picture with Caption
[Layout 9] Title and Vertical Text
[Layout 10] Vertical Title and Text
```

この番号を控えておく。自社テンプレートは並び順が異なる場合がある。

## Step 2a: Python スクリプトで使う

確認したインデックスをスクリプト先頭の定数として設定する。`demo_builder.py` がそのまま参考になる。

```python
from pptx_components import *

# テンプレートを使う場合はパスを指定する
# None にすると python-pptx のデフォルト（25.4 x 19.05cm 標準サイズ）を使用
TEMPLATE = "demo_template.pptx"   # make_template.py で生成

# pptx_decode.py template.pptx layout で確認した番号を設定する
LAYOUT_TITLE   = 0    # Title Slide       (CENTER_TITLE + SUBTITLE)
LAYOUT_CONTENT = 1    # Title and Content (TITLE + BODY)
LAYOUT_BLANK   = 6    # Blank

# create the ppt object
prs = create_ppt(TEMPLATE)

# タイトルスライド
add_title_slide(prs, LAYOUT_TITLE, "タイトル", "サブタイトル")

# コンテンツスライド（全シェイプを手動配置）
slide = add_blank_slide(prs, LAYOUT_BLANK)
add_content_header(slide, "スライドタイトル", "サブタイトル", font_name="Meiryo UI")

# セッションタイトルスライド
add_section_slide(prs, LAYOUT_TITLE, "タイトル", "サブタイトル")

# save the file 
save_ppt(prs, "output.pptx")
```

> **NOTE:** `add_blank_slide()` はレイアウトのシェイプをすべて削除した空スライドを返す。テンプレートの背景・ロゴ・底バーはマスターレイヤーに存在するため、`add_blank_slide()` を使っても引き継がれる。

## Step 2b: YAML + yaml2pptx.py で使う

YAMLファイルの先頭に `template:` キーと、各スライドの `layout_index:` を記述する。

```yaml
template: ./demo_template.pptx   # make_template.py で生成したテンプレート

slides:
  - index: 0
    layout_index: 0        # Step 1 で確認した番号（Title Slide）
    shapes:
      - fn: set_placeholder
        placeholder_type: CENTER_TITLE
        placeholder_idx: 0
        text: タイトル
      - fn: set_placeholder
        placeholder_type: SUBTITLE
        placeholder_idx: 1
        text: サブタイトル

  - index: 1
    layout_index: 6        # Blank
    shapes:
      - fn: add_text
        x: 1.5
        y: 0.5
        w: 26.0
        h: 1.2
        text: スライドタイトル
        size: 24.0
        bold: true
        color: '#000000'
        font_name: Meiryo UI
```

実行：

```bash
yaml2pptx.py slides.yaml output.pptx
```

`--template` オプションを指定するとYAML内の `template:` キーより優先される：

```bash
yaml2pptx.py slides.yaml output.pptx --template other_template.pptx
```

## テンプレートの優先順位（yaml2pptx）

| 優先度 | 指定方法 |
|---|---|
| 高 | `--template` オプション |
| 中 | YAMLファイルの `template:` キー |
| 低（省略） | テンプレートなし（python-pptxデフォルト） |

## テンプレートなしで使う

`TEMPLATE = None` または `create_ppt()` を引数なしで呼ぶと、python-pptxのデフォルトプレゼンテーション（**25.4 × 19.05 cm、標準4:3サイズ**）を使用する。`demo_builder.py` がこのパターンになっている。

`make_template.py` で生成した `demo_template.pptx` はワイドスクリーン（33.867 × 19.05 cm）のため、スライドサイズを揃えたい場合はテンプレートを指定する方が良い。

デフォルト（テンプレートなし）のレイアウトインデックス：

| インデックス | 名前 | 主なプレースホルダー |
|---|---|---|
| 0 | Title Slide | CENTER\_TITLE + SUBTITLE |
| 1 | Title and Content | TITLE + BODY |
| 2 | Section Header | TITLE + BODY |
| 5 | Title Only | TITLE |
| 6 | Blank | なし |

```python
prs = create_ppt()              # テンプレートなし（25.4 x 19.05cm）
slide = add_blank_slide(prs, 6) # Blankレイアウト
```
