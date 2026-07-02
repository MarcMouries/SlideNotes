import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import xml2js from "xml2js";

const { parseStringPromise: parseXml, Builder: XmlBuilder } = xml2js;

export const NO_TITLE = "No title";

const TITLE_PLACEHOLDER_TYPES = new Set(["ctrTitle", "title"]);
const BODY_PLACEHOLDER_TYPES = new Set(["body", "obj", "subTitle"]);
const NOTES_BODY_PLACEHOLDER_TYPES = new Set(["body"]);
const NOTES_SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const NOTES_MASTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const xmlBuilder = new XmlBuilder({ renderOpts: { pretty: false } });

function asArray(value) {
    if (value === undefined || value === null) {
        return [];
    }

    return Array.isArray(value) ? value : [value];
}

function textNodeValue(node) {
    if (typeof node === "string") {
        return node.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    }

    if (node && typeof node._ === "string") {
        return node._.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    }

    return "";
}

function extractRunText(run) {
    return asArray(run?.["a:t"]).map(textNodeValue).join("");
}

function extractParagraphText(paragraph) {
    const runs = [
        ...asArray(paragraph?.["a:r"]),
        ...asArray(paragraph?.["a:fld"]),
    ];

    return runs.map(extractRunText).join("");
}

function extractShapeText(shape) {
    const paragraphs = asArray(shape?.["p:txBody"]?.[0]?.["a:p"]);
    const text = paragraphs.map(extractParagraphText).join("\n").trim();

    return text || NO_TITLE;
}

function titlePlaceholderType(shape) {
    return shape?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.$?.type;
}

function placeholderType(shape) {
    return shape?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.$?.type;
}

function slideNumberFromPath(fileName) {
    const match = fileName.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    return match ? Number(match[1]) : null;
}

function slideRelationshipPath(slideFile) {
    const slideName = path.posix.basename(slideFile);

    return `ppt/slides/_rels/${slideName}.rels`;
}

function sourcePartPathFromRelationshipPath(relsFile) {
    const relsDir = path.posix.dirname(relsFile);
    const sourceDir = path.posix.dirname(relsDir);
    const sourceFile = path.posix.basename(relsFile, ".rels");

    return path.posix.join(sourceDir, sourceFile);
}

function relationshipTargetToZipPath(sourceFile, target) {
    if (!target) {
        return null;
    }

    if (target.startsWith("/")) {
        return target.slice(1);
    }

    const sourcePart = sourcePartPathFromRelationshipPath(sourceFile);

    return path.posix.normalize(path.posix.join(path.posix.dirname(sourcePart), target));
}

async function parseZipXml(zip, fileName) {
    const file = zip.files[fileName];

    if (!file) {
        return null;
    }

    return parseXml(await file.async("string"));
}

function writeZipXml(zip, fileName, xmlObject) {
    zip.file(fileName, xmlBuilder.buildObject(xmlObject));
}

function reportProgress(onProgress, update) {
    if (typeof onProgress === "function") {
        onProgress(update);
    }
}

async function removeRelationshipTypes(zip, relsFile, relationshipTypes) {
    const relsXml = await parseZipXml(zip, relsFile);

    if (!relsXml?.Relationships) {
        return;
    }

    const relationships = asArray(relsXml.Relationships.Relationship);
    const keptRelationships = relationships.filter((relationship) => {
        return !relationshipTypes.has(relationship?.$?.Type);
    });

    relsXml.Relationships.Relationship = keptRelationships;
    writeZipXml(zip, relsFile, relsXml);
}

async function removeNotesContentTypes(zip) {
    const contentTypesXml = await parseZipXml(zip, "[Content_Types].xml");

    if (!contentTypesXml?.Types) {
        return;
    }

    const overrides = asArray(contentTypesXml.Types.Override);

    contentTypesXml.Types.Override = overrides.filter((override) => {
        const partName = override?.$?.PartName || "";

        return !partName.startsWith("/ppt/notesSlides/")
            && !partName.startsWith("/ppt/notesMasters/");
    });

    writeZipXml(zip, "[Content_Types].xml", contentTypesXml);
}

export default class SlideNotes {
    constructor() {
        this.slides = [];
    }

    async load(pptxPath, options = {}) {
        const { onProgress } = options;

        reportProgress(onProgress, {
            phase: "read-file",
            current: 0,
            total: 1,
        });

        const data = await readFile(pptxPath);

        reportProgress(onProgress, {
            phase: "read-file",
            current: 1,
            total: 1,
        });

        const zip = await JSZip.loadAsync(data);
        const slideFiles = Object.keys(zip.files)
            .filter((fileName) => slideNumberFromPath(fileName) !== null)
            .sort((left, right) => slideNumberFromPath(left) - slideNumberFromPath(right));

        this.slides = [];

        reportProgress(onProgress, {
            phase: "read-slides",
            current: 0,
            total: slideFiles.length,
        });

        for (let index = 0; index < slideFiles.length; index++) {
            const slideFile = slideFiles[index];
            const content = await zip.files[slideFile].async("string");
            const result = await parseXml(content);
            const title = this.extractTitle(result);
            const notesFile = await this.findNotesFileForSlide(zip, slideFile);
            const notes = notesFile
                ? this.extractNotes(await parseXml(await zip.files[notesFile].async("string")))
                : "";

            this.slides.push({
                getTitle: () => title,
                getNotes: () => notes,
            });

            reportProgress(onProgress, {
                phase: "read-slides",
                current: index + 1,
                total: slideFiles.length,
            });
        }

        return this;
    }

    extractTitle(slideXml) {
        const shapes = asArray(slideXml?.["p:sld"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"]);

        for (const shape of shapes) {
            if (TITLE_PLACEHOLDER_TYPES.has(titlePlaceholderType(shape))) {
                return extractShapeText(shape);
            }
        }

        const fallbackTitle = shapes
            .filter((shape) => !BODY_PLACEHOLDER_TYPES.has(placeholderType(shape)))
            .map(extractShapeText)
            .find((text) => text !== NO_TITLE);

        if (fallbackTitle) {
            return fallbackTitle;
        }

        return NO_TITLE;
    }

    extractNotes(notesXml) {
        const shapes = asArray(notesXml?.["p:notes"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"]);
        const bodyShapes = shapes.filter((shape) => NOTES_BODY_PLACEHOLDER_TYPES.has(placeholderType(shape)));
        const bodyTexts = bodyShapes
            .map(extractShapeText)
            .filter((text) => text !== NO_TITLE);

        if (bodyShapes.length > 0) {
            return bodyTexts.join("\n\n");
        }

        return shapes
            .filter((shape) => !placeholderType(shape))
            .map(extractShapeText)
            .filter((text) => text !== NO_TITLE)
            .join("\n\n");
    }

    async findNotesFileForSlide(zip, slideFile) {
        const relsFile = slideRelationshipPath(slideFile);
        const relsXml = await parseZipXml(zip, relsFile);
        const notesRelationship = asArray(relsXml?.Relationships?.Relationship)
            .find((relationship) => relationship?.$?.Type === NOTES_SLIDE_REL_TYPE);

        if (!notesRelationship) {
            return null;
        }

        const notesFile = relationshipTargetToZipPath(relsFile, notesRelationship.$.Target);

        return zip.files[notesFile] ? notesFile : null;
    }

    async exportSlideNotesText(pptxPath, outputPath, options = {}) {
        const { onProgress } = options;
        const pres = await this.load(pptxPath, { onProgress });
        const blocks = [];

        reportProgress(onProgress, {
            phase: "export-notes",
            current: 0,
            total: pres.slides.length,
        });

        for (let index = 0; index < pres.slides.length; index++) {
            const slide = pres.slides[index];

            blocks.push(`Slide #${index + 1}: ${slide.getTitle()}\n${slide.getNotes()}`);
            reportProgress(onProgress, {
                phase: "export-notes",
                current: index + 1,
                total: pres.slides.length,
            });
        }

        await writeFile(outputPath, `${blocks.join("\n\n")}\n`, "utf8");

        return outputPath;
    }

    async removeNotes(pptxPath, outputPath) {
        const data = await readFile(pptxPath);
        const zip = await JSZip.loadAsync(data);
        const noteFilePrefixes = ["ppt/notesSlides/", "ppt/notesMasters/"];
        const relationshipTypesToRemove = new Set([
            NOTES_SLIDE_REL_TYPE,
            NOTES_MASTER_REL_TYPE,
        ]);

        for (const fileName of Object.keys(zip.files)) {
            if (noteFilePrefixes.some((prefix) => fileName.startsWith(prefix))) {
                zip.remove(fileName);
            }
        }

        const relFiles = Object.keys(zip.files).filter((fileName) => fileName.endsWith(".rels"));

        for (const relsFile of relFiles) {
            await removeRelationshipTypes(zip, relsFile, relationshipTypesToRemove);
        }

        await removeNotesContentTypes(zip);
        await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));

        return outputPath;
    }
}
