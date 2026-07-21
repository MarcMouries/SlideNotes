import { describe, expect, test } from "bun:test";
import { mkdtemp, readFile, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import JSZip from "jszip";
import xml2js from "xml2js";
import SlideNotes from "../src/index.js";
import {
    ZIP_METHOD_DEFLATED,
    ZIP_METHOD_STORED,
    danglingRelationshipIds,
    unresolvedRelationshipTargets,
    zipCompressionMethods,
} from "./helpers/package-integrity.js";

const { parseStringPromise: parseXml } = xml2js;

const fixturePath = path.resolve("test/fixture/simple-with-notes.pptx");

async function tempFile(fileName) {
    const tempDir = await mkdtemp(path.join(tmpdir(), "slidenotes-"));

    return path.join(tempDir, fileName);
}

async function writeDeflatedCopy(pptxPath) {
    const zip = await JSZip.loadAsync(await readFile(pptxPath));
    const outputPath = await tempFile("deflated-input.pptx");

    await writeFile(outputPath, await zip.generateAsync({
        type: "nodebuffer",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
    }));

    return outputPath;
}

describe("remove-notes output integrity", () => {
    test("does not grow the file when the input is compressed", async () => {
        const inputPath = await writeDeflatedCopy(fixturePath);
        const outputPath = await tempFile("without-notes.pptx");

        await new SlideNotes().removeNotes(inputPath, outputPath);

        const inputSize = (await stat(inputPath)).size;
        const outputSize = (await stat(outputPath)).size;

        expect(outputSize).toBeLessThanOrEqual(inputSize);
    });

    test("compresses every content entry in the output package", async () => {
        const outputPath = await tempFile("without-notes.pptx");

        await new SlideNotes().removeNotes(fixturePath, outputPath);

        const methods = zipCompressionMethods(await readFile(outputPath));

        expect(methods.size).toBeGreaterThan(0);

        for (const [name, { method, uncompressedSize }] of methods) {
            if (uncompressedSize === 0) {
                expect([ZIP_METHOD_STORED, ZIP_METHOD_DEFLATED]).toContain(method);
            } else {
                expect({ name, method }).toEqual({ name, method: ZIP_METHOD_DEFLATED });
            }
        }
    });

    test("removes the notes master reference from presentation.xml", async () => {
        const outputPath = await tempFile("without-notes.pptx");

        await new SlideNotes().removeNotes(fixturePath, outputPath);

        const zip = await JSZip.loadAsync(await readFile(outputPath));
        const presentationXml = await zip.files["ppt/presentation.xml"].async("string");

        expect(presentationXml).not.toContain("notesMasterIdLst");
        expect(presentationXml).not.toContain("notesMasterId");
    });

    test("leaves no dangling relationship references in presentation.xml", async () => {
        const outputPath = await tempFile("without-notes.pptx");

        await new SlideNotes().removeNotes(fixturePath, outputPath);

        const zip = await JSZip.loadAsync(await readFile(outputPath));
        const presentationXml = await zip.files["ppt/presentation.xml"].async("string");
        const relsXml = await parseXml(await zip.files["ppt/_rels/presentation.xml.rels"].async("string"));
        const relationships = relsXml.Relationships.Relationship || [];

        expect(relationships.length).toBeGreaterThan(0);
        expect(await danglingRelationshipIds(zip, "ppt/presentation.xml")).toEqual([]);
    });

    test("every relationship target in the output resolves to an existing part", async () => {
        const outputPath = await tempFile("without-notes.pptx");

        await new SlideNotes().removeNotes(fixturePath, outputPath);

        const zip = await JSZip.loadAsync(await readFile(outputPath));

        expect(await unresolvedRelationshipTargets(zip)).toEqual([]);
    });

    test("the output still loads with the same slides and titles", async () => {
        const outputPath = await tempFile("without-notes.pptx");
        const before = await new SlideNotes().load(fixturePath);

        await new SlideNotes().removeNotes(fixturePath, outputPath);

        const after = await new SlideNotes().load(outputPath);

        expect(after.slides.map((slide) => slide.getTitle()))
            .toEqual(before.slides.map((slide) => slide.getTitle()));
        expect(after.slides.every((slide) => slide.getNotes() === "")).toBe(true);
    });
});
