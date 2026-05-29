# 白根福音教会 会計システム 仕様書 — ビルド設定

TOPLEVEL=chapter

MD_FILES := $(shell find . -maxdepth 1 -name "*.md" | sort)
TEX_FILES := $(MD_FILES:.md=.tex)

# メインターゲット
all: $(TEX_FILES)
	lualatex main.tex
	@if [ -s main.idx ]; then \
		upmendex main.idx; \
		lualatex main.tex; \
	else \
		echo "main.idx is empty, skipping upmendex"; \
	fi

# MarkdownからTeXへの変換
%.tex: %.md
	md2tex $< -o $@

# 生成物とすべての中間ファイルを削除
clean:
	rm -f $(TEX_FILES)
	find . -name "*.aux" -delete
	find . -name "*.log" -delete
	find . -name "*.toc" -delete
	find . -name "*.out" -delete
	find . -name "*.idx" -delete
	find . -name "*.ilg" -delete
	find . -name "*.ind" -delete

.PHONY: all clean
