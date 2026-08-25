PY := .venv/bin/python
SRC ?= input/source.pptx
TPL ?= input/template.pptx
OUT ?= output/converted.pptx

help:
	@echo "targets: inspect-src inspect-tpl fonts extract apply render-src render-out qa clean"

inspect-src: ; $(PY) tools/inspect_deck.py $(SRC)
inspect-tpl: ; $(PY) tools/inspect_deck.py $(TPL)
fonts:       ; bash tools/install-fonts.sh $(SRC) $(TPL)
extract:     ; $(PY) tools/extract.py $(SRC) work
apply:       ; $(PY) tools/apply.py $(TPL) work/content.json mapping.json $(OUT)
render-src:  ; bash tools/render.sh $(SRC) output/render-src
render-out:  ; bash tools/render.sh $(OUT) output/render-out
qa:          ; $(PY) tools/qa.py output/render-src output/render-out output/qa

clean: ; rm -rf work output/render-* output/qa

.PHONY: help inspect-src inspect-tpl fonts extract apply render-src render-out qa clean
