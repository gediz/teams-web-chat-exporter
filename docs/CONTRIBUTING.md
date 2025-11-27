# Contributing

Thank you for your interest in contributing. This project uses the WXT Framework to support Chrome, Edge, and Firefox.

## Getting Started

Please refer to [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) for setup, installation, and testing instructions.

## Documentation

- [README.md](README.md): User guide and features.
- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md): Technical design and component communication.
- [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md): Build and testing instructions.
- [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md): How to build, package, and publish to stores.
- [docs/MIGRATION_NOTES.md](docs/MIGRATION_NOTES.md): History of the migration to WXT.

## Code Style

- **Framework**: Svelte 5 (Popup UI) + TypeScript.
- **Linting**: Run `npm run check` to verify types.
- **I18N**: Add new locales in `src/i18n/locales/`.

## Submitting Changes

1. Fork the repository.
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Make your changes.
4. Test in both Chrome and Firefox.
5. Run `npm run check` to verify types.
6. Run `npm run build && npm run build:firefox` to ensure builds work.
7. Commit with descriptive message: `feat: add export cancellation`
8. Submit a Pull Request with:
   - Description of the change
   - Testing performed
   - Screenshots (if UI changes)
