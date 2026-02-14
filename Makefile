.PHONY: test test-unit test-integ lint clean release tag help

# Test environment
export XDG_CACHE_HOME := /tmp
export XDG_STATE_HOME := /tmp
export XDG_DATA_HOME := /tmp

# Version detection
CURRENT_VERSION := $(shell git describe --tags --abbrev=0 2>/dev/null || echo "v0.0.0")
NEXT_VERSION := $(shell echo $(CURRENT_VERSION) | awk -F. '{$$NF = $$NF + 1;} 1' | sed 's/ /./g')

test: test-unit test-integ

test-unit:
	@nvim --headless -u NONE -i NONE -l test/test_docxedit.lua

test-integ:
	@nvim --headless -u NONE -i NONE -l test/integration_test.lua

test-watch:
	@command -v entr >/dev/null 2>&1 || (echo "entr not found. Install: brew install entr" && exit 1)
	@find lua test -name "*.lua" | entr -c $(MAKE) test

lint: lint-lua lint-style lint-doc

lint-lua:
	@if command -v luacheck >/dev/null 2>&1; then \
		luacheck lua/ --codes; \
	else \
		echo "luacheck not found (optional). Install: luarocks install luacheck"; \
	fi

lint-style:
	@if command -v stylua >/dev/null 2>&1; then \
		stylua --check lua/ test/; \
	else \
		echo "stylua not found (optional). Install: cargo install stylua"; \
	fi

lint-doc:
	@test -f README.md || (echo "README.md missing" && exit 1)
	@test -f doc/docxedit.txt || (echo "doc/docxedit.txt missing" && exit 1)

fixup:
	@if command -v stylua >/dev/null 2>&1; then \
		stylua lua/ test/; \
	else \
		echo "stylua not found. Install: cargo install stylua"; \
		exit 1; \
	fi

tag:
	@echo "Current version: $(CURRENT_VERSION)"
	@echo "Next version: $(NEXT_VERSION)"
	@read -p "Enter version (default: $(NEXT_VERSION)): " VERSION; \
	VERSION=$${VERSION:-$(NEXT_VERSION)}; \
	if ! echo "$$VERSION" | grep -qE '^v[0-9]+\.[0-9]+\.[0-9]+$$'; then \
		echo "Version must be in format v1.2.3"; \
		exit 1; \
	fi; \
	$(MAKE) test || exit 1; \
	git tag -a "$$VERSION" -m "Release $$VERSION"; \
	echo "Tag $$VERSION created. Push with: git push origin $$VERSION"

release:
	@BRANCH=$$(git rev-parse --abbrev-ref HEAD); \
	if [ "$$BRANCH" != "main" ]; then \
		echo "Must be on main branch (currently on $$BRANCH)"; \
		exit 1; \
	fi
	@if ! git diff-index --quiet HEAD --; then \
		echo "Uncommitted changes detected"; \
		exit 1; \
	fi
	@git pull --rebase || exit 1
	@$(MAKE) test || exit 1
	@echo "Current version: $(CURRENT_VERSION)"
	@echo "Next version: $(NEXT_VERSION)"
	@read -p "Enter version (default: $(NEXT_VERSION)): " VERSION; \
	VERSION=$${VERSION:-$(NEXT_VERSION)}; \
	if ! echo "$$VERSION" | grep -qE '^v[0-9]+\.[0-9]+\.[0-9]+$$'; then \
		echo "Version must be in format v1.2.3"; \
		exit 1; \
	fi; \
	if git rev-parse "$$VERSION" >/dev/null 2>&1; then \
		echo "Tag $$VERSION already exists"; \
		exit 1; \
	fi; \
	read -p "Create and push $$VERSION? (y/N): " CONFIRM; \
	if [ "$$CONFIRM" != "y" ] && [ "$$CONFIRM" != "Y" ]; then \
		echo "Cancelled"; \
		exit 1; \
	fi; \
	git tag -a "$$VERSION" -m "Release $$VERSION"; \
	git push origin "$$VERSION"; \
	echo "Released $$VERSION"

clean:
	@rm -rf /tmp/docxedit_*
	@rm -rf /tmp/nvim-*
	@rm -f .luacheckcache
	@find . -name "*.tmp" -delete 2>/dev/null || true
	@find . -name ".DS_Store" -delete 2>/dev/null || true

clean-cache:
	@rm -rf ~/.cache/nvim/docxedit/ 2>/dev/null || true
	@rm -rf /tmp/xdg-cache/docxedit/ 2>/dev/null || true

health:
	@nvim -c "checkhealth docxedit" -c "quitall"

check-deps:
	@echo "Required:"
	@command -v nvim >/dev/null 2>&1 && echo "  [✓] nvim" || echo "  [✗] nvim"
	@command -v git >/dev/null 2>&1 && echo "  [✓] git" || echo "  [✗] git"
	@echo "Optional:"
	@command -v luacheck >/dev/null 2>&1 && echo "  [✓] luacheck" || echo "  [✗] luacheck"
	@command -v stylua >/dev/null 2>&1 && echo "  [✓] stylua" || echo "  [✗] stylua"
	@command -v entr >/dev/null 2>&1 && echo "  [✓] entr" || echo "  [✗] entr"

install-hooks:
	@echo '#!/bin/sh' > .git/hooks/pre-commit
	@echo 'make test || exit 1' >> .git/hooks/pre-commit
	@chmod +x .git/hooks/pre-commit
	@echo "Pre-commit hook installed"

info:
	@echo "Project: nvim-docx"
	@echo "Version: $(CURRENT_VERSION)"
	@echo "Branch: $$(git rev-parse --abbrev-ref HEAD)"
	@echo "Commit: $$(git rev-parse --short HEAD)"
	@echo "Remote: $$(git remote get-url origin 2>/dev/null || echo 'none')"

help:
	@echo "Available targets:"
	@echo "  test         - Run all tests"
	@echo "  test-unit    - Run unit tests"
	@echo "  test-integ   - Run integration tests"
	@echo "  test-watch   - Run tests on file changes"
	@echo "  lint         - Run all linting"
	@echo "  lint-lua     - Check Lua syntax"
	@echo "  lint-style   - Check code style"
	@echo "  lint-doc     - Check documentation"
	@echo "  fixup        - Auto-fix code style"
	@echo "  tag          - Create version tag"
	@echo "  release      - Create and push release"
	@echo "  clean        - Clean temp files"
	@echo "  clean-cache  - Clean nvim cache"
	@echo "  health       - Run health check"
	@echo "  check-deps   - Check dependencies"
	@echo "  install-hooks- Install git hooks"
	@echo "  info         - Show project info"
	@echo "  help         - Show this help"
