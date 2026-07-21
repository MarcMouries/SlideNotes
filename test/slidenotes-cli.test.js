import { describe, expect, test } from "bun:test";
import { spawnSync } from "node:child_process";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { writePptxFixture } from "./helpers/pptx-fixture.js";

const projectRoot = fileURLToPath(new URL("..", import.meta.url));

function derivedOutputPath(filePath, suffix, extension) {
    const parsedPath = path.parse(filePath);

    return path.join(parsedPath.dir, `${parsedPath.name}${suffix}${extension}`);
}

describe("slidenotes CLI", () => {
    test("prints slide count and titles", async () => {
        const filePath = await writePptxFixture([
            { number: 1, textRuns: ["Intro"] },
            { number: 2, textRuns: ["Details"] },
        ]);

        const result = spawnSync("bun", ["bin/slidenotes.js", "list", filePath], {
            cwd: projectRoot,
            encoding: "utf8",
        });

        expect(result.status).toBe(0);
        expect(result.stderr).toBe("");
        expect(result.stdout).toBe("Total slides: 2\n\n1: Intro\n2: Details\n\nNotes on 0 of 2 slides.\n");
    });

    test("marks slides that have notes with a word count", async () => {
        const filePath = await writePptxFixture([
            { number: 1, textRuns: ["Intro"], notes: [["Say hello to everyone."]] },
            { number: 2, textRuns: ["Details"] },
            { number: 3, textRuns: ["Solo"], notes: [["Pause."]] },
        ]);

        const result = spawnSync("bun", ["bin/slidenotes.js", "list", filePath], {
            cwd: projectRoot,
            encoding: "utf8",
        });

        expect(result.status).toBe(0);
        expect(result.stdout).toBe([
            "Total slides: 3",
            "",
            "1: Intro — notes: 4 words",
            "2: Details",
            "3: Solo — notes: 1 word",
            "",
            "Notes on 2 of 3 slides.",
            "",
        ].join("\n"));
    });

    test("prints usage when a pptx path is missing", () => {
        const result = spawnSync("bun", ["bin/slidenotes.js"], {
            cwd: projectRoot,
            encoding: "utf8",
        });

        expect(result.status).toBe(1);
        expect(result.stdout).toBe("");
        expect(result.stderr).toContain("Usage:");
        expect(result.stderr).toContain("slidenotes list <pptx file path>");
        expect(result.stderr).toContain("slidenotes export-notes <pptx file path>");
        expect(result.stderr).toContain("slidenotes import-notes <pptx file path> [notes text file]");
        expect(result.stderr).toContain("slidenotes remove-notes <pptx file path>");
    });

    test("exports notes through the CLI to a derived text path", async () => {
        const filePath = await writePptxFixture([
            { number: 1, textRuns: ["Intro"], notes: [["Say hello."]] },
        ]);
        const outputPath = derivedOutputPath(filePath, "-speaker-notes", ".txt");

        const result = spawnSync("bun", ["bin/slidenotes.js", "export-notes", filePath], {
            cwd: projectRoot,
            encoding: "utf8",
        });

        expect(result.status).toBe(0);
        expect(result.stdout).toContain(`Wrote slide notes to ${outputPath}`);
        expect(await readFile(outputPath, "utf8")).toBe("Slide #1: Intro\nSay hello.\n");
    });

    test("imports notes through the CLI with an explicit notes file", async () => {
        const filePath = await writePptxFixture([
            { number: 1, textRuns: ["Intro"], notes: [["Old note."]] },
        ]);
        const notesPath = derivedOutputPath(filePath, "-edited", ".txt");
        const outputPath = derivedOutputPath(filePath, "-with-notes", ".pptx");

        await Bun.write(notesPath, "Slide #1: Intro\nNew note.\n");

        const result = spawnSync("bun", ["bin/slidenotes.js", "import-notes", filePath, notesPath], {
            cwd: projectRoot,
            encoding: "utf8",
        });

        expect(result.status).toBe(0);
        expect(result.stdout).toContain(`Wrote presentation with notes to ${outputPath}`);

        const listResult = spawnSync("bun", ["bin/slidenotes.js", "list", outputPath], {
            cwd: projectRoot,
            encoding: "utf8",
        });

        expect(listResult.stdout).toContain("1: Intro — notes: 2 words");
    });

    test("imports notes from the default exported path when none is given", async () => {
        const filePath = await writePptxFixture([
            { number: 1, textRuns: ["Intro"], notes: [["Original note."]] },
        ]);

        const exportResult = spawnSync("bun", ["bin/slidenotes.js", "export-notes", filePath], {
            cwd: projectRoot,
            encoding: "utf8",
        });

        expect(exportResult.status).toBe(0);

        const importResult = spawnSync("bun", ["bin/slidenotes.js", "import-notes", filePath], {
            cwd: projectRoot,
            encoding: "utf8",
        });

        expect(importResult.status).toBe(0);
        expect(importResult.stdout).toContain(derivedOutputPath(filePath, "-with-notes", ".pptx"));
    });

    test("fails with a clear error when the notes file is missing", async () => {
        const filePath = await writePptxFixture([
            { number: 1, textRuns: ["Intro"], notes: [["A note."]] },
        ]);
        const missingPath = derivedOutputPath(filePath, "-does-not-exist", ".txt");

        const result = spawnSync("bun", ["bin/slidenotes.js", "import-notes", filePath, missingPath], {
            cwd: projectRoot,
            encoding: "utf8",
        });

        expect(result.status).toBe(1);
        expect(result.stderr).toContain(`Notes file not found: ${missingPath}`);
    });

    test("removes notes through the CLI to a derived pptx path", async () => {
        const filePath = await writePptxFixture([
            { number: 1, textRuns: ["Intro"], notes: [["Private note."]] },
        ]);
        const outputPath = derivedOutputPath(filePath, "-without-notes", ".pptx");

        const result = spawnSync("bun", ["bin/slidenotes.js", "remove-notes", filePath], {
            cwd: projectRoot,
            encoding: "utf8",
        });

        expect(result.status).toBe(0);
        expect(result.stdout).toContain(`Wrote presentation without notes to ${outputPath}`);

        const zip = await JSZip.loadAsync(await readFile(outputPath));

        expect(Object.keys(zip.files).some((fileName) => fileName.startsWith("ppt/notesSlides/"))).toBe(false);
    });
});
