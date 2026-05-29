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

## Installation

```bash
# Clone the repository
git clone https://github.com/akirayan/pptx_decode.git
cd pptx_decode

# Install dependencies
pip install -r requirements.txt
```

To use the library from your own project, either copy the `.py` files into your project folder, or add the cloned directory to `sys.path`:

```python
import sys
sys.path.insert(0, "/path/to/pptx_decode")
from pptx_components import *
```

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

## 背景・開発動機

PowerPointを使うたびに、同じことで苦労してきた。

フォントサイズの調整、シェイプの位置合わせ、色の設定。GUIの操作は直感的に見えて、実際には細かい数値を思い通りに制御できない。操作するたびに「もっと楽にできないか」と感じていた。

転機はAIとの対話だった。AIがPPTXを生成するPythonスクリプトを提示したとき、「これは使える」と直感した。スクリプトを自分で書けば、座標も色も数値で完全に制御できる。GUIのもどかしさから解放される。

ただし、まずPPTXの内部構造を理解しなければならない。そこで最初に作ったのが `pptx_decode.py` だ。既存のPPTXファイルを読み込んで、構造を人間が読める形で出力する。

構造が理解できたら次は生成だ。最初はPythonスクリプトにスライドの内容を直接書いた。しかし、資料を更新するたびにコードを修正するのは手間がかかる。そこで発想を変えた。**スライドの内容はYAMLで書く。生成はCLIで行う。** `pptx2yaml.py` でPPTXの内容をYAMLに書き出し、`yaml2pptx.py` でYAMLからPPTXを生成する。

### Blank Slide をキャンバスとして使う

コーディングで最も理解に苦しんだのは、スライドマスターのレイアウトに埋め込まれた **Placeholder** というシェイプだった。レイアウトからの継承関係が複雑で、思い通りに動かない場面が何度もあった。

解決策は「逃げること」だった。**Blank Slideを作り、全シェイプを削除する**。残るのは何もない白紙のスライドだ。そこに `add_rect()`、`add_text()`、`add_line()` で自由に描く。テンプレートのロゴや底バーはスライドマスターレイヤーに存在するため、全シェイプを削除しても継承される。

このアプローチは、三十数年前のSunOS XViewプログラミング経験と共鳴している。Xlibの描画関数を直接呼び出してPanel上に自由に描いた、あの発想と同じだ。フレームワークの「良い使い方」から外れているかもしれないが、自分が何をしているかが明確で、制御が完全に自分の手にある。

## License

MIT
