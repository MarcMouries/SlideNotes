import { describe, expect, test } from "bun:test";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import JSZip from "jszip";
import xml2js from "xml2js";
import SlideNotes, { parseSpeakerNotesText } from "../src/index.js";
import {
    duplicateRelationshipIds,
    unresolvedRelationshipTargets,
    zipCompressionMethods,
    ZIP_METHOD_DEFLATED,
} from "./helpers/package-integrity.js";

const { parseStringPromise: parseXml } = xml2js;

const fixturePath = path.resolve("test/fixture/simple-with-notes.pptx");

async function tempDir() {
    return mkdtemp(path.join(tmpdir(), "slidenotes-"));
}

async function tempFile(fileName) {
    return path.join(await tempDir(), fileName);
}

async function writeNotesFile(text) {
    const notesPath = await tempFile("notes.txt");

    await writeFile(notesPath, text, "utf8");

    return notesPath;
}

async function loadZip(pptxPath) {
    return JSZip.loadAsync(await readFile(pptxPath));
}

async function importNotes(pptxPath, notesText) {
    const notesPath = await writeNotesFile(notesText);
    const outputPath = await tempFile("with-notes.pptx");

    await new SlideNotes().importSlideNotesText(pptxPath, notesPath, outputPath);

    return outputPath;
}

async function assertPackageIntegrity(pptxPath) {
    const zip = await loadZip(pptxPath);

    expect(await unresolvedRelationshipTargets(zip)).toEqual([]);
    expect(await duplicateRelationshipIds(zip)).toEqual([]);

    const methods = zipCompressionMethods(await readFile(pptxPath));

    for (const [name, { method, uncompressedSize }] of methods) {
        if (uncompressedSize > 0) {
            expect({ name, method }).toEqual({ name, method: ZIP_METHOD_DEFLATED });
        }
    }
}

describe("parseSpeakerNotesText", () => {
    test("parses the export format including empty blocks", () => {
        const parsed = parseSpeakerNotesText([
            "Slide #1: Opening",
            "Welcome everyone.",
            "Here is the plan.",
            "",
            "Slide #2: Budget",
            "",
            "",
            "Slide #3: Appendix",
            "Final note.",
            "",
        ].join("\n"));

        expect([...parsed.entries()]).toEqual([
            [1, "Welcome everyone.\nHere is the plan."],
            [2, ""],
            [3, "Final note."],
        ]);
    });

    test("preserves blank lines inside a notes block", () => {
        const parsed = parseSpeakerNotesText("Slide #1: Title\nFirst paragraph.\n\nSecond paragraph.\n");

        expect(parsed.get(1)).toBe("First paragraph.\n\nSecond paragraph.");
    });

    test("normalizes CRLF line endings", () => {
        const parsed = parseSpeakerNotesText("Slide #1: Title\r\nLine one.\r\nLine two.\r\n");

        expect(parsed.get(1)).toBe("Line one.\nLine two.");
    });

    test("handles a missing trailing newline", () => {
        const parsed = parseSpeakerNotesText("Slide #1: Title\nOnly note");

        expect(parsed.get(1)).toBe("Only note");
    });

    test("treats a notes line that looks like a header as a header (documented limitation)", () => {
        const parsed = parseSpeakerNotesText("Slide #1: Title\nSlide #7: this is really a note\n");

        expect(parsed.get(1)).toBe("");
        expect(parsed.has(7)).toBe(true);
    });

    test("rejects duplicate slide numbers", () => {
        expect(() => parseSpeakerNotesText("Slide #1: A\nx\n\nSlide #1: B\ny\n"))
            .toThrow("Slide #1 more than once");
    });

    test("rejects text with no headers", () => {
        expect(() => parseSpeakerNotesText("just some text\n")).toThrow("no \"Slide #N:\" headers");
    });

    test("rejects content before the first header", () => {
        expect(() => parseSpeakerNotesText("preamble\nSlide #1: Title\nnote\n"))
            .toThrow("before the first");
    });
});

describe("import-notes into slides with existing notes parts", () => {
    test("replaces and clears notes per the file, leaving unlisted slides untouched", async () => {
        const notesText = [
            "Slide #1: Opening",
            "Rewritten opening notes.",
            "With a second line.",
            "",
            "Slide #2: Budget",
            "",
        ].join("\n");

        const beforeZip = await loadZip(fixturePath);
        const untouchedBefore = await beforeZip.files["ppt/notesSlides/notesSlide3.xml"].async("string");

        const outputPath = await importNotes(fixturePath, notesText);
        const reader = await new SlideNotes().load(outputPath);

        expect(reader.slides[0].getNotes()).toBe("Rewritten opening notes.\nWith a second line.");
        expect(reader.slides[1].getNotes()).toBe("");
        expect(reader.slides[2].getNotes()).toBe("This note belongs to a slide without a title.");

        const afterZip = await loadZip(outputPath);
        const untouchedAfter = await afterZip.files["ppt/notesSlides/notesSlide3.xml"].async("string");

        expect(untouchedAfter).toBe(untouchedBefore);
        await assertPackageIntegrity(outputPath);
    });

    test("preserves blank-line paragraph breaks through a load round-trip", async () => {
        const outputPath = await importNotes(fixturePath, "Slide #1: Opening\nFirst.\n\nSecond.\n");
        const reader = await new SlideNotes().load(outputPath);

        expect(reader.slides[0].getNotes()).toBe("First.\n\nSecond.");
    });

    test("escapes XML special characters in notes text", async () => {
        const outputPath = await importNotes(fixturePath, "Slide #1: Opening\nBudget < revenue & costs > 0\n");
        const reader = await new SlideNotes().load(outputPath);

        expect(reader.slides[0].getNotes()).toBe("Budget < revenue & costs > 0");
        await assertPackageIntegrity(outputPath);
    });
});

describe("import-notes creating notes slides", () => {
    // Manufactures the "deck has a notes master but this slide has no notes
    // slide" state, which pptxgenjs decks never exhibit naturally.
    async function writeFixtureWithoutNotesSlide3(stripSlideRelsFile = false) {
        const zip = await loadZip(fixturePath);

        zip.remove("ppt/notesSlides/notesSlide3.xml");
        zip.remove("ppt/notesSlides/_rels/notesSlide3.xml.rels");

        const contentTypes = await parseXml(await zip.files["[Content_Types].xml"].async("string"));

        contentTypes.Types.Override = contentTypes.Types.Override
            .filter((override) => override.$.PartName !== "/ppt/notesSlides/notesSlide3.xml");
        zip.file("[Content_Types].xml", new xml2js.Builder({ renderOpts: { pretty: false } }).buildObject(contentTypes));

        if (stripSlideRelsFile) {
            zip.remove("ppt/slides/_rels/slide3.xml.rels");
        } else {
            const slideRels = await parseXml(await zip.files["ppt/slides/_rels/slide3.xml.rels"].async("string"));

            slideRels.Relationships.Relationship = slideRels.Relationships.Relationship
                .filter((relationship) => !relationship.$.Type.endsWith("/notesSlide"));
            zip.file("ppt/slides/_rels/slide3.xml.rels", new xml2js.Builder({ renderOpts: { pretty: false } }).buildObject(slideRels));
        }

        const modifiedPath = await tempFile("missing-notesslide3.pptx");

        await writeFile(modifiedPath, await zip.generateAsync({ type: "nodebuffer" }));

        return modifiedPath;
    }

    test("creates a new notes slide part with correct rels and content type", async () => {
        const inputPath = await writeFixtureWithoutNotesSlide3();
        const outputPath = await importNotes(inputPath, "Slide #3: No title\nBrand new note.\n");
        const zip = await loadZip(outputPath);

        expect(zip.files["ppt/notesSlides/notesSlide5.xml"]).toBeDefined();

        const notesRels = await parseXml(await zip.files["ppt/notesSlides/_rels/notesSlide5.xml.rels"].async("string"));
        const targets = notesRels.Relationships.Relationship.map((relationship) => relationship.$.Target);

        expect(targets).toContain("../notesMasters/notesMaster1.xml");
        expect(targets).toContain("../slides/slide3.xml");

        const contentTypes = await zip.files["[Content_Types].xml"].async("string");

        expect(contentTypes).toContain("/ppt/notesSlides/notesSlide5.xml");

        const reader = await new SlideNotes().load(outputPath);

        expect(reader.slides[2].getNotes()).toBe("Brand new note.");
        await assertPackageIntegrity(outputPath);
    });

    test("creates the slide rels file when it is missing entirely", async () => {
        const inputPath = await writeFixtureWithoutNotesSlide3(true);
        const outputPath = await importNotes(inputPath, "Slide #3: No title\nNote for orphan slide.\n");
        const reader = await new SlideNotes().load(outputPath);

        expect(reader.slides[2].getNotes()).toBe("Note for orphan slide.");
        await assertPackageIntegrity(outputPath);
    });
});

describe("import-notes into a deck without a notes master", () => {
    async function removedNotesDeck() {
        const outputPath = await tempFile("without-notes.pptx");

        await new SlideNotes().removeNotes(fixturePath, outputPath);

        return outputPath;
    }

    test("recreates the notes master and inserts notesMasterIdLst in schema order", async () => {
        const strippedPath = await removedNotesDeck();
        const outputPath = await importNotes(strippedPath, "Slide #1: Opening\nRestored note.\n");
        const zip = await loadZip(outputPath);

        expect(zip.files["ppt/notesMasters/notesMaster1.xml"]).toBeDefined();

        // The notes master must get its own theme copy, not share the slide
        // master's: PowerPoint repairs decks with a shared master theme.
        const masterRels = await parseXml(await zip.files["ppt/notesMasters/_rels/notesMaster1.xml.rels"].async("string"));

        expect(masterRels.Relationships.Relationship[0].$.Target).toBe("../theme/theme2.xml");
        expect(zip.files["ppt/theme/theme2.xml"]).toBeDefined();

        const contentTypes = await zip.files["[Content_Types].xml"].async("string");

        expect(contentTypes).toContain("/ppt/theme/theme2.xml");

        const presentationXml = await zip.files["ppt/presentation.xml"].async("string");
        const masterIndex = presentationXml.indexOf("<p:sldMasterIdLst");
        const notesMasterIndex = presentationXml.indexOf("<p:notesMasterIdLst");
        const slideListIndex = presentationXml.indexOf("<p:sldIdLst");

        expect(masterIndex).toBeGreaterThan(-1);
        expect(notesMasterIndex).toBeGreaterThan(masterIndex);
        expect(slideListIndex).toBeGreaterThan(notesMasterIndex);

        const reader = await new SlideNotes().load(outputPath);

        expect(reader.slides[0].getNotes()).toBe("Restored note.");
        await assertPackageIntegrity(outputPath);
    });

    test("creates the notes master only once for multiple slides", async () => {
        const strippedPath = await removedNotesDeck();
        const outputPath = await importNotes(strippedPath, [
            "Slide #1: Opening",
            "Note one.",
            "",
            "Slide #2: Budget",
            "Note two.",
            "",
        ].join("\n"));
        const zip = await loadZip(outputPath);
        const masterParts = Object.keys(zip.files)
            .filter((fileName) => /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(fileName));

        expect(masterParts).toEqual(["ppt/notesMasters/notesMaster1.xml"]);

        const reader = await new SlideNotes().load(outputPath);

        expect(reader.slides[0].getNotes()).toBe("Note one.");
        expect(reader.slides[1].getNotes()).toBe("Note two.");
        await assertPackageIntegrity(outputPath);
    });
});

describe("import-notes round-trips", () => {
    test("export -> remove-notes -> import -> export reproduces the original text", async () => {
        const reader = new SlideNotes();
        const originalTextPath = await tempFile("original.txt");

        await reader.exportSlideNotesText(fixturePath, originalTextPath);

        const strippedPath = await tempFile("stripped.pptx");

        await new SlideNotes().removeNotes(fixturePath, strippedPath);

        const restoredPath = await tempFile("restored.pptx");

        await new SlideNotes().importSlideNotesText(strippedPath, originalTextPath, restoredPath);

        const roundTripTextPath = await tempFile("roundtrip.txt");

        await new SlideNotes().exportSlideNotesText(restoredPath, roundTripTextPath);

        expect(await readFile(roundTripTextPath, "utf8")).toBe(await readFile(originalTextPath, "utf8"));
        await assertPackageIntegrity(restoredPath);
    });

    test("export -> import into the original -> export is stable", async () => {
        const originalTextPath = await tempFile("original.txt");

        await new SlideNotes().exportSlideNotesText(fixturePath, originalTextPath);

        const reimportedPath = await tempFile("reimported.pptx");

        await new SlideNotes().importSlideNotesText(fixturePath, originalTextPath, reimportedPath);

        const roundTripTextPath = await tempFile("roundtrip.txt");

        await new SlideNotes().exportSlideNotesText(reimportedPath, roundTripTextPath);

        expect(await readFile(roundTripTextPath, "utf8")).toBe(await readFile(originalTextPath, "utf8"));
    });
});

describe("import-notes errors", () => {
    test("rejects slide numbers not present in the deck without writing output", async () => {
        const notesPath = await writeNotesFile("Slide #9: Ghost\nBoo.\n");
        const outputPath = await tempFile("never-written.pptx");

        expect(new SlideNotes().importSlideNotesText(fixturePath, notesPath, outputPath))
            .rejects.toThrow("Slide #9 not found in presentation (4 slides)");
        expect(existsSync(outputPath)).toBe(false);
    });

    test("reports a missing notes file by path", async () => {
        const missingPath = path.join(await tempDir(), "nope.txt");
        const outputPath = await tempFile("never-written.pptx");

        expect(new SlideNotes().importSlideNotesText(fixturePath, missingPath, outputPath))
            .rejects.toThrow(`Notes file not found: ${missingPath}`);
    });
});
