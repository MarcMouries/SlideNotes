import { describe, expect, test } from "bun:test";
import { mkdir, readFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import SlideNotes, { NO_TITLE } from "../src/index.js";
import { demoSlides, writeDemoPresentationWithNotes } from "./helpers/demo-presentation.js";

const artifactsDir = path.resolve("test-output");
const withNotesPath = path.join(artifactsDir, "simple-with-notes.pptx");
const exportedNotesPath = path.join(artifactsDir, "simple-notes.txt");
const withoutNotesPath = path.join(artifactsDir, "simple-without-notes.pptx");

async function zipFileNames(filePath) {
    const zip = await JSZip.loadAsync(await readFile(filePath));

    return Object.keys(zip.files);
}

describe("demo artifacts", () => {
    test("writes visible files for the notes workflow", async () => {
        await mkdir(artifactsDir, { recursive: true });

        await writeDemoPresentationWithNotes(withNotesPath);
        await new SlideNotes().exportSlideNotesText(withNotesPath, exportedNotesPath);
        await new SlideNotes().removeNotes(withNotesPath, withoutNotesPath);

        expect(await readFile(exportedNotesPath, "utf8")).toBe([
            `Slide #1: ${demoSlides[0].title}`,
            demoSlides[0].notes,
            "",
            `Slide #2: ${demoSlides[1].title}`,
            demoSlides[1].notes,
            "",
            `Slide #3: ${NO_TITLE}`,
            demoSlides[2].notes,
            "",
            `Slide #4: ${demoSlides[3].title}`,
            demoSlides[3].notes,
            "",
        ].join("\n"));

        const withNotesFiles = await zipFileNames(withNotesPath);
        const withoutNotesFiles = await zipFileNames(withoutNotesPath);
        const reader = await new SlideNotes().load(withoutNotesPath);

        expect(withNotesFiles.some((fileName) => fileName.startsWith("ppt/notesSlides/"))).toBe(true);
        expect(withoutNotesFiles.some((fileName) => fileName.startsWith("ppt/notesSlides/"))).toBe(false);
        expect(reader.slides).toHaveLength(demoSlides.length);
        expect(reader.slides.map((slide) => slide.getTitle())).toEqual(demoSlides.map((slide) => {
            return slide.expectedTitle || slide.title;
        }));
    });
});
