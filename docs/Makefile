SPHINXOPTS    ?=
SPHINXBUILD   ?= sphinx-build
SOURCEDIR     = source
BUILDDIR      = build

.PHONY: help clean html open

help:
	@$(SPHINXBUILD) -M help "$(SOURCEDIR)" "$(BUILDDIR)" $(SPHINXOPTS) $(O)

clean:
	rm -rf $(BUILDDIR)

html:
	$(SPHINXBUILD) -b html -W --keep-going $(SPHINXOPTS) $(SOURCEDIR) $(BUILDDIR)/html
	@echo "\nDoc générée → $(BUILDDIR)/html/index.html"

open: html
	xdg-open $(BUILDDIR)/html/index.html 2>/dev/null || open $(BUILDDIR)/html/index.html
