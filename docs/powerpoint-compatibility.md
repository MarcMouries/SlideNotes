# PowerPoint Compatibility Notes

This document records two rounds of PowerPoint compatibility bugs in SlideNotes, their root causes, and the fixes. It exists because none of the standard validation tools can catch these problems — the constraints live only inside PowerPoint's parser — and because the fixes look arbitrary without this context.

## Bug 1: `remove-notes` produced larger files that PowerPoint offered to repair

**Symptoms** (reported July 2026): running `remove-notes` on a deck produced an output *larger* than the input despite removing content, and PowerPoint prompted *"PowerPoint found a problem with content… can attempt to repair"* when opening it.

### Cause A: uncompressed output

JSZip's `generateAsync` defaults to `STORE` (no compression). The output package was written entirely uncompressed, while real `.pptx` files are deflate-compressed — so a 21 KB deck became a 45 KB deck even with the notes parts deleted.

**Fix:** every package write goes through `writeZipToFile()` in `src/SlideNotes.js`, which passes `compression: "DEFLATE"` explicitly. The regression test `test/remove-notes.test.js` reads compression methods straight out of the zip central directory (JSZip does not expose them after loading) and asserts every content entry is deflated.

### Cause B: dangling notes-master reference

`remove-notes` deleted the `ppt/notesMasters/` parts and their relationships but left this in `ppt/presentation.xml`:

```xml
<p:notesMasterIdLst><p:notesMasterId r:id="rId6"/></p:notesMasterIdLst>
```

with `rId6` no longer present in `presentation.xml.rels`. A dangling `r:id` reference is a classic repair-prompt trigger.

**Fix:** `removeNotesMasterIdList()` removes the element alongside the parts. The test suite asserts no `r:id` in `presentation.xml` lacks a matching relationship (`test/helpers/package-integrity.js`).

## Bug 2: `import-notes` output repaired when it had to create a notes master

**Symptom:** importing notes into a deck whose notes infrastructure had been fully stripped (e.g. by our own `remove-notes`) produced a file PowerPoint offered to repair. Importing into a deck that still had notes parts was fine.

### Investigation

The file was valid by every measurable standard: it passed ECMA-376 schema validation (`pml.xsd` via xmllint), OPC relationship checks, python-pptx, and Apache POI's strict parser. In fact it was *more* standards-compliant than files PowerPoint happily opens. The cause was found empirically, by generating a battery of `.pptx` files each isolating one transformation and opening each in PowerPoint:

| Probe | Isolated change | PowerPoint |
|---|---|---|
| JSZip repack, no content change | zip layer | opens |
| Every `.rels` re-serialized via xml2js | rels round-trip | opens |
| `[Content_Types].xml` re-serialized | content-types round-trip | opens |
| `presentation.xml` re-serialized, order unchanged | XML round-trip | opens |
| `notesMasterIdLst` moved to schema position | element order | **repair** |
| Generated notes-slide XML shape | new part shape | opens |
| Stale `<Notes>` count in `docProps/app.xml` | metadata mismatch | irrelevant |
| Notes slide created under existing master | creation machinery | opens |
| Schema position **plus** own cloned theme | interaction probe | opens |

### Root cause: an interaction, not a single defect

Two conditions, each individually accepted by PowerPoint, trigger the repair prompt **only in combination**:

1. The notes master's relationships point at the **same theme part** as the slide master (`theme1.xml`). pptxgenjs builds decks this way; genuine PowerPoint decks always give each master its own theme part.
2. `p:notesMasterIdLst` sits in its **schema-correct position** in `presentation.xml` (between `sldMasterIdLst` and `sldIdLst`) — which is also where genuine PowerPoint files put it.

pptxgenjs decks survive condition 1 only because pptxgenjs also violates the schema by writing `notesMasterIdLst` *after* `sldIdLst`. Our import code inserted the element in the schema-correct position (verified against real PowerPoint-authored files) while inheriting the shared-theme layout from the deck — producing the one combination PowerPoint rejects.

### Fix

`ensureNotesMaster()` in `src/SlideNotes.js` clones the presentation's theme into a new `ppt/theme/themeN.xml` (with its own content-type override) and points the newly created notes master at the clone. This matches how PowerPoint itself structures decks and keeps the schema-correct element order. The state-(c) test in `test/import-notes.test.js` asserts the cloned theme part, its content-type registration, and the element order.

## Lessons for future OOXML work here

- **Schema/OPC validity is necessary but not sufficient.** PowerPoint enforces constraints no schema expresses, and tolerates violations schemas reject. The only authoritative check is opening the file in PowerPoint.
- **Test the creation paths, not just the mutation paths.** Editing existing parts in place is far safer than synthesizing parts; every repair prompt came from a path that created or removed package structure.
- **Bisect empirically.** When PowerPoint rejects a generated file, produce variants that each isolate one transformation and decode the pass/fail matrix (see `scratchpad` battery scripts from July 2026; pattern documented here so it can be recreated).
- **Copy structures from genuine PowerPoint files**, not from other generators — pptxgenjs's output deviates from PowerPoint's own conventions in ways that happen to work only in pptxgenjs's exact configuration.
