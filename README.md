# SlideNotes

A tool for working with PowerPoint presentations.

It can pull speaker notes out of a `.pptx` file into a readable text document, and it can create a copy of the deck with all speaker notes removed.

It is useful when you want to review the talk track separately, share slides without private notes, or quickly check what is inside a presentation without opening PowerPoint.

## Using The CLI

Make `slidenotes` available on your PATH:

```bash
bun install
bun link
```

After this, run the CLI directly as `slidenotes`.

Print the slide list:

```bash
slidenotes list path/to/input.pptx
```

Export slide titles and speaker notes to a text file:

```bash
slidenotes export-notes path/to/input.pptx
```

For `path/to/input.pptx`, this writes:

```text
path/to/input-speaker-notes.txt
```

Generate a copy of a presentation without speaker notes:

```bash
slidenotes remove-notes path/to/input.pptx
```

For `path/to/input.pptx`, this writes:

```text
path/to/input-without-notes.pptx
```

When run in an interactive terminal, `slidenotes` shows progress while reading slides and exporting notes. Progress is written to stderr so stdout remains usable for scripts.

## Notes Export Format

The exported text file uses one block per slide:

```text
Slide #1: Opening
Welcome everyone.
Here is the plan.

Slide #2: Budget
Mention Q4 numbers.
Pause for questions.

Slide #3: No title
This note belongs to a slide without a title.
```

Slides without a detected title are exported as `No title`.

## Test Fixture

Generate a simple PowerPoint fixture:

```bash
bun run generate:fixture
```

This writes `test/fixture/simple-with-notes.pptx`.

Try it with:

```bash
slidenotes list ./test/fixture/simple-with-notes.pptx
slidenotes export-notes ./test/fixture/simple-with-notes.pptx
slidenotes remove-notes ./test/fixture/simple-with-notes.pptx
```

## Developing

Install dependencies:

```bash
bun install
```

Run tests:

```bash
bun test
```

Build a standalone executable for your current platform:

```bash
bun run build
```

The local build writes `dist/slidenotes`.

Build release binaries for macOS and Windows:

```bash
bun run build:release
```

This writes:

- `dist/slidenotes-macos-arm64`
- `dist/slidenotes-macos-x64`
- `dist/slidenotes-windows-x64.exe`

## Releasing

GitHub releases are created from version tags.

1. Update the version in `package.json`.
2. Commit and push the change.
3. Create and push a version tag:

```bash
git tag v0.1.0
git push origin v0.1.0
```

Pushing the tag runs `.github/workflows/release.yml`. The workflow runs tests, builds macOS and Windows CLI binaries, packages them, and attaches them to a GitHub Release.

Release artifacts:

- `slidenotes-macos-arm64.tar.gz` for macOS Apple Silicon
- `slidenotes-macos-x64.tar.gz` for macOS Intel
- `slidenotes-windows-x64.zip` for Windows x64

## Project Layout

- `src/SlideNotes.js` - core PPTX reader, notes exporter, and notes remover
- `src/index.js` - public export
- `bin/slidenotes.js` - command-line entry point
- `scripts/generate-test-fixture.js` - fixture generator
- `test/` - tests and generated fixture helpers
