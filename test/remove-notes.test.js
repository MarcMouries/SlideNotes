import { describe, expect, test } from "bun:test";
import { mkdtemp, readFile, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import JSZip from "jszip";
import xml2js from "xml2js";
import SlideNotes from "../src/index.js";

const { parseStringPromise: parseXml } = xml2js;

const fixturePath = path.resolve("test/fixture/simple-with-notes.pptx");

const ZIP_END_OF_CENTRAL_DIR_SIGNATURE = Buffer.from([0x50, 0x4b, 0x05, 0x06]);
const ZIP_METHOD_STORED = 0;
const ZIP_METHOD_DEFLATED = 8;

async function tempFile(fileName) {
    const tempDir = await mkdtemp(path.join(tmpdir(), "slidenotes-"));

    return path.join(tempDir, fileName);
}

// Reads each central directory entry's compression method straight from the
// zip bytes, since JSZip does not expose them after loading.
function zipCompressionMethods(buffer) {
    const endOfCentralDir = buffer.lastIndexOf(ZIP_END_OF_CENTRAL_DIR_SIGNATURE);
    const entryCount = buffer.readUInt16LE(endOfCentralDir + 10);
    const methods = new Map();
    let offset = buffer.readUInt32LE(endOfCentralDir + 16);

    for (let index = 0; index < entryCount; index++) {
        const method = buffer.readUInt16LE(offset + 10);
        const uncompressedSize = buffer.readUInt32LE(offset + 24);
        const nameLength = buffer.readUInt16LE(offset + 28);
        const extraLength = buffer.readUInt16LE(offset + 30);
        const commentLength = buffer.readUInt16LE(offset + 32);
        const name = buffer.toString("utf8", offset + 46, offset + 46 + nameLength);

        methods.set(name, { method, uncompressedSize });
        offset += 46 + nameLength + extraLength + commentLength;
    }

    return methods;
}

async function writeDeflatedCopy(pptxPath) {
    const zip = await JSZip.loadAsync(await readFile(pptxPath));
    const outputPath = await tempFile("deflated-input.pptx");
    const { writeFile } = await import("node:fs/promises");

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
        const relationshipIds = new Set(relationships.map((relationship) => relationship.$.Id));
        const referencedIds = [...presentationXml.matchAll(/r:(?:id|embed)="([^"]+)"/g)]
            .map((match) => match[1]);

        expect(referencedIds.length).toBeGreaterThan(0);

        for (const referencedId of referencedIds) {
            expect(relationshipIds.has(referencedId)).toBe(true);
        }
    });

    test("every relationship target in the output resolves to an existing part", async () => {
        const outputPath = await tempFile("without-notes.pptx");

        await new SlideNotes().removeNotes(fixturePath, outputPath);

        const zip = await JSZip.loadAsync(await readFile(outputPath));
        const relsFiles = Object.keys(zip.files).filter((fileName) => fileName.endsWith(".rels"));

        expect(relsFiles.length).toBeGreaterThan(0);

        for (const relsFile of relsFiles) {
            const relsXml = await parseXml(await zip.files[relsFile].async("string"));
            const relationships = relsXml.Relationships.Relationship || [];

            for (const relationship of relationships) {
                if (relationship.$.TargetMode === "External") {
                    continue;
                }

                const relsDir = path.posix.dirname(relsFile);
                const sourceDir = path.posix.dirname(relsDir);
                const target = relationship.$.Target;
                const targetPath = target.startsWith("/")
                    ? target.slice(1)
                    : path.posix.normalize(path.posix.join(sourceDir, target));

                expect({ relsFile, target, exists: Boolean(zip.files[targetPath]) })
                    .toEqual({ relsFile, target, exists: true });
            }
        }
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
