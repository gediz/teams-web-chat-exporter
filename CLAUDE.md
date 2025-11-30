# CLAUDE.md - AI Assistant Guide

This file provides AI assistants with context about the Teams Chat Exporter codebase.

## Project Overview

**Teams Chat Exporter** is a cross-browser extension (Chrome, Edge, Firefox) that exports Microsoft Teams web chat conversations to JSON, CSV, HTML, or text formats.

- **Version**: 1.1.0
- **Framework**: WXT (Vite-based)
- **Languages**: TypeScript + Svelte 5
- **Supported Browsers**: Chrome, Edge, Firefox
- **UI Languages**: 14 (en, zh-CN, pt-BR, nl, fr, de, it, ja, ko, ru, es, tr, ar, he)

## Quick Reference

### Key Files
- **Popup**: [src/entrypoints/popup/App.svelte](src/entrypoints/popup/App.svelte)
- **Background**: [src/entrypoints/background.ts](src/entrypoints/background.ts)
- **Content**: [src/entrypoints/content.ts](src/entrypoints/content.ts)
- **Config**: [wxt.config.ts](wxt.config.ts)

### Build Commands
See [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) for all build and testing commands.

### Architecture
See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for:
- Project structure
- Component communication
- Message flow
- Data structures
- Browser compatibility

### Code Conventions
See [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md) for:
- Naming conventions
- TypeScript guidelines
- Svelte patterns
- Common development tasks

## Documentation

- [README.md](README.md): User guide
- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md): Technical design
- [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md): Code conventions and common tasks
- [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md): Build and testing
- [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md): Store publishing
- [docs/MANUAL_INSTALL.md](docs/MANUAL_INSTALL.md): Installation steps
- [docs/TODO.md](docs/TODO.md): Project roadmap

## When Helping Users

### For Code Questions
1. Check [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for technical design
2. Check [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md) for code conventions and common tasks
3. Reference specific files with line numbers (e.g., `App.svelte:42`)

### For Build Issues
1. Check [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) for all build commands
2. Common fix: `rm -rf .output .wxt node_modules && npm install`

### For Feature Requests
1. Check [docs/TODO.md](docs/TODO.md) for planned features
2. Consider cross-browser compatibility
3. Add i18n keys for new UI strings (14 languages)

---

**Last Updated**: 2025-11-30
**Version**: 1.1.0
