# Contributing to GriMoire

Thank you for considering a contribution. This guide covers the basics.

## How to contribute

1. Fork the repository
2. Create a feature branch (`git checkout -b my-feature`)
3. Make your changes
4. Run the build and tests
5. Commit and push to your fork
6. Open a pull request against `main`

## Development setup

### Web part

```bash
cd grimoire-webpart
npm install
npx heft build --clean
```

### Backend

```bash
cd grimoire-backend
npm install
func start
```

### Documentation site

The docs site lives in a separate repo: [grimoire-hie.github.io](https://github.com/grimoire-hie/grimoire-hie.github.io).

## Code style

- TypeScript with ESLint (`@microsoft/eslint-config-spfx`)
- SPFx build targets ES5 — no `for...of` on Map/Set, no RegExp `u` flag
- All promises must be awaited or voided (`@typescript-eslint/no-floating-promises`)
- Use `logService` for logging, not `console.log`

## Testing

```bash
cd grimoire-webpart
npx heft test --clean
```

Run a single test file:

```bash
npx heft test --clean -- --testPathPattern="McpResultMapper"
```

## Pull requests

- Keep PRs focused on a single change
- Include a clear description of what changed and why
- Make sure the build passes (`npx heft build --clean`)
- Add or update tests when changing behavior

## Issues

Use [GitHub Issues](https://github.com/grimoire-hie/grimoire-spfx/issues) to report bugs or suggest features. Include steps to reproduce for bugs.

## License

By contributing, you agree that your contributions will be licensed under the [MIT License](LICENSE).
