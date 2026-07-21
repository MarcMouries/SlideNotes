# SlideNotes

A tool for working with PowerPoint presentations.

It can pull speaker notes out of a `.pptx` file into a readable text document, apply an edited notes file back onto a deck, and create a copy of the deck with all speaker notes removed.

It is useful when you want to review or edit the talk track separately, share slides without private notes, or quickly check what is inside a presentation without opening PowerPoint.

## Using The CLI

Make `slidenotes` available on your PATH:

```bash
bun install
bun link
```

After this, run the CLI directly as `slidenotes`.

Print the slide list, marking slides that have speaker notes:

```bash
slidenotes list path/to/input.pptx
```

```text
Total slides: 4

1: Opening — notes: 8 words
2: Budget — notes: 6 words
3: No title — notes: 9 words
4: Appendix

Notes on 3 of 4 slides.
```

Export slide titles and speaker notes to a text file:

```bash
slidenotes export-notes path/to/input.pptx
```

For `path/to/input.pptx`, this writes:

```text
path/to/input-speaker-notes.txt
```

Apply a notes text file back onto a presentation (the round trip of `export-notes`):

```bash
slidenotes import-notes path/to/input.pptx path/to/notes.txt
```

The notes file argument is optional and defaults to `path/to/input-speaker-notes.txt`, so `export-notes`, edit, `import-notes` works without extra arguments. For `path/to/input.pptx`, this writes:

```text
path/to/input-with-notes.pptx
```

For every `Slide #N` block in the file, the slide's notes are replaced with the block's text; an empty block clears that slide's notes. Slides not listed in the file are left untouched. Importing works even on decks whose notes parts were fully stripped by `remove-notes`. One caveat: a notes line that itself starts with `Slide #N:` is indistinguishable from a slide header, so avoid that pattern in notes text.

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
slidenotes import-notes ./test/fixture/simple-with-notes.pptx
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

## PowerPoint Compatibility

SlideNotes rewrites `.pptx` packages, and PowerPoint enforces undocumented constraints that no schema validator catches. [docs/powerpoint-compatibility.md](docs/powerpoint-compatibility.md) records the repair-prompt bugs we hit, their root causes (including a subtle theme-sharing × element-order interaction), and how to debug this class of problem. Read it before changing any code that creates or removes package parts.

## Project Layout

- `src/SlideNotes.js` - core PPTX reader, notes exporter/importer, and notes remover
- `src/notes-master-template.js` - known-good notes master XML used when importing into a deck without one
- `src/index.js` - public export
- `bin/slidenotes.js` - command-line entry point
- `scripts/generate-test-fixture.js` - fixture generator
- `test/` - tests and generated fixture helpers
