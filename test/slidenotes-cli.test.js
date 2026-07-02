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
        expect(result.stdout).toBe("Total slides: 2\n\n1: Intro\n2: Details\n");
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
