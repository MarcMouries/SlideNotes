import { describe, expect, test } from "bun:test";
import { mkdtemp, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import JSZip from "jszip";
import SlideNotes from "../src/index.js";
import { writePptxFixture } from "./helpers/pptx-fixture.js";

async function tempFile(fileName) {
    const tempDir = await mkdtemp(path.join(tmpdir(), "slidenotes-"));

    return path.join(tempDir, fileName);
}

describe("notes workflows", () => {
    test("exports slide titles and notes to a text file", async () => {
        const pptxPath = await writePptxFixture([
            {
                number: 1,
                textRuns: ["Opening"],
                notes: [["Welcome ", "everyone."], ["Here is the plan."]],
            },
            {
                number: 2,
                textRuns: ["Budget"],
                notes: [["Mention Q4 numbers."]],
            },
            {
                number: 3,
                textRuns: ["Appendix"],
            },
        ]);
        const outputPath = await tempFile("talktrack.txt");

        await new SlideNotes().exportSlideNotesText(pptxPath, outputPath);

        expect(await readFile(outputPath, "utf8")).toBe([
            "Slide #1: Opening",
            "Welcome everyone.\nHere is the plan.",
            "",
            "Slide #2: Budget",
            "Mention Q4 numbers.",
            "",
            "Slide #3: Appendix",
            "",
            "",
        ].join("\n"));
    });

    test("loads an empty notes string for slides without speaker notes", async () => {
        const pptxPath = await writePptxFixture([
            {
                number: 1,
                textRuns: ["With notes"],
                notes: [["Presenter guidance."]],
            },
            {
                number: 2,
                textRuns: ["No notes"],
            },
        ]);

        const reader = await new SlideNotes().load(pptxPath);

        expect(reader.slides[0].getNotes()).toBe("Presenter guidance.");
        expect(reader.slides[1].getTitle()).toBe("No notes");
        expect(reader.slides[1].getNotes()).toBe("");
    });

    test("writes a copy of a presentation without speaker notes", async () => {
        const pptxPath = await writePptxFixture([
            {
                number: 1,
                textRuns: ["Opening"],
                notes: [["Private talking point."]],
            },
            {
                number: 2,
                textRuns: ["No notes here"],
            },
        ]);
        const outputPath = await tempFile("without-notes.pptx");

        await new SlideNotes().removeNotes(pptxPath, outputPath);

        const zip = await JSZip.loadAsync(await readFile(outputPath));
        const fileNames = Object.keys(zip.files);

        expect(fileNames.some((fileName) => fileName.startsWith("ppt/notesSlides/"))).toBe(false);
        expect(fileNames.some((fileName) => fileName.startsWith("ppt/notesMasters/"))).toBe(false);

        const slideRelationships = await zip.files["ppt/slides/_rels/slide1.xml.rels"].async("string");
        const contentTypes = await zip.files["[Content_Types].xml"].async("string");

        expect(slideRelationships).not.toContain("notesSlide");
        expect(contentTypes).not.toContain("notesSlides");
    });
});
