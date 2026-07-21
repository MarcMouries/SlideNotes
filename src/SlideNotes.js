import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import xml2js from "xml2js";
import NOTES_MASTER_XML from "./notes-master-template.js";

const { parseStringPromise: parseXml, Builder: XmlBuilder } = xml2js;

export const NO_TITLE = "No title";

const TITLE_PLACEHOLDER_TYPES = new Set(["ctrTitle", "title"]);
const BODY_PLACEHOLDER_TYPES = new Set(["body", "obj", "subTitle"]);
const NOTES_BODY_PLACEHOLDER_TYPES = new Set(["body"]);
const NOTES_SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const NOTES_MASTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const THEME_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const RELATIONSHIPS_XMLNS = "http://schemas.openxmlformats.org/package/2006/relationships";
const NOTES_SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const NOTES_MASTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";
const THEME_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml";
const DRAWING_XMLNS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const RELATIONSHIPS_DOC_XMLNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PRESENTATION_XMLNS = "http://schemas.openxmlformats.org/presentationml/2006/main";

// Child element order required by the CT_Presentation schema; inserting a
// child anywhere else can make PowerPoint offer to repair the file.
const PRESENTATION_CHILD_ORDER = [
    "p:sldMasterIdLst",
    "p:notesMasterIdLst",
    "p:handoutMasterIdLst",
    "p:sldIdLst",
    "p:sldSz",
    "p:notesSz",
    "p:smartTags",
    "p:embeddedFontLst",
    "p:custShowLst",
    "p:photoAlbum",
    "p:custDataLst",
    "p:kinsoku",
    "p:defaultTextStyle",
    "p:modifyVerifier",
    "p:extLst",
];

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

function sortedSlideFiles(zip) {
    return Object.keys(zip.files)
        .filter((fileName) => slideNumberFromPath(fileName) !== null)
        .sort((left, right) => slideNumberFromPath(left) - slideNumberFromPath(right));
}

async function writeZipToFile(zip, outputPath) {
    await writeFile(outputPath, await zip.generateAsync({
        type: "nodebuffer",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
    }));
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

function nextRelationshipId(relsXml) {
    const relationships = asArray(relsXml?.Relationships?.Relationship);
    const maxId = relationships.reduce((max, relationship) => {
        const match = relationship?.$?.Id?.match(/^rId(\d+)$/);

        return match ? Math.max(max, Number(match[1])) : max;
    }, 0);

    return `rId${maxId + 1}`;
}

async function ensureRelsXml(zip, relsFile) {
    const relsXml = await parseZipXml(zip, relsFile);

    if (relsXml?.Relationships) {
        return relsXml;
    }

    return { Relationships: { $: { xmlns: RELATIONSHIPS_XMLNS } } };
}

function appendRelationship(relsXml, { id, type, target }) {
    const relationships = asArray(relsXml.Relationships.Relationship);

    relationships.push({ $: { Id: id, Type: type, Target: target } });
    relsXml.Relationships.Relationship = relationships;
}

function findRelationshipByType(relsXml, type) {
    return asArray(relsXml?.Relationships?.Relationship)
        .find((relationship) => relationship?.$?.Type === type) || null;
}

function nextPartNumber(zip, partPattern) {
    const maxNumber = Object.keys(zip.files).reduce((max, fileName) => {
        const match = fileName.match(partPattern);

        return match ? Math.max(max, Number(match[1])) : max;
    }, 0);

    return maxNumber + 1;
}

async function addContentTypeOverride(zip, partName, contentType) {
    const contentTypesXml = await parseZipXml(zip, "[Content_Types].xml");

    if (!contentTypesXml?.Types) {
        throw new Error("Invalid presentation: missing [Content_Types].xml");
    }

    const overrides = asArray(contentTypesXml.Types.Override);

    if (overrides.some((override) => override?.$?.PartName === partName)) {
        return;
    }

    overrides.push({ $: { PartName: partName, ContentType: contentType } });
    contentTypesXml.Types.Override = overrides;
    writeZipXml(zip, "[Content_Types].xml", contentTypesXml);
}

function notesParagraphsFromText(text) {
    if (text === "") {
        return [{}];
    }

    return text.split("\n").map((line) => {
        return line === "" ? {} : { "a:r": [{ "a:t": [line] }] };
    });
}

function notesBodyShape(paragraphs) {
    return {
        "p:nvSpPr": [{
            "p:cNvPr": [{ $: { id: "3", name: "Notes Placeholder" } }],
            "p:cNvSpPr": [{ "a:spLocks": [{ $: { noGrp: "1" } }] }],
            "p:nvPr": [{ "p:ph": [{ $: { type: "body", idx: "1" } }] }],
        }],
        "p:spPr": [{}],
        "p:txBody": [{
            "a:bodyPr": [{}],
            "a:lstStyle": [{}],
            "a:p": paragraphs,
        }],
    };
}

function setNotesBodyParagraphs(notesXml, paragraphs) {
    const spTree = notesXml?.["p:notes"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0];

    if (!spTree) {
        throw new Error("Invalid notes slide: missing shape tree");
    }

    const shapes = asArray(spTree["p:sp"]);
    const bodyShape = shapes.find((shape) => NOTES_BODY_PLACEHOLDER_TYPES.has(placeholderType(shape)));

    if (!bodyShape) {
        shapes.push(notesBodyShape(paragraphs));
        spTree["p:sp"] = shapes;
        return;
    }

    const txBody = bodyShape["p:txBody"]?.[0];

    if (txBody) {
        txBody["a:p"] = paragraphs;
    } else {
        bodyShape["p:txBody"] = [{
            "a:bodyPr": [{}],
            "a:lstStyle": [{}],
            "a:p": paragraphs,
        }];
    }
}

// Rebuilds the p:presentation element with the new child in schema position.
// A plain property assignment would append it after existing keys, and
// xml2js serializes children in key order.
function insertPresentationChildInSchemaOrder(presentationElement, key, value) {
    const keyIndex = PRESENTATION_CHILD_ORDER.indexOf(key);
    const rebuilt = {};
    let inserted = false;

    for (const [existingKey, existingValue] of Object.entries(presentationElement)) {
        const existingIndex = PRESENTATION_CHILD_ORDER.indexOf(existingKey);

        if (!inserted && existingIndex !== -1 && existingIndex > keyIndex) {
            rebuilt[key] = value;
            inserted = true;
        }

        rebuilt[existingKey] = existingValue;
    }

    if (!inserted) {
        rebuilt[key] = value;
    }

    return rebuilt;
}

function buildNotesSlideXml(paragraphs, slideNumber) {
    return {
        "p:notes": {
            $: {
                "xmlns:a": DRAWING_XMLNS,
                "xmlns:r": RELATIONSHIPS_DOC_XMLNS,
                "xmlns:p": PRESENTATION_XMLNS,
            },
            "p:cSld": [{
                "p:spTree": [{
                    "p:nvGrpSpPr": [{
                        "p:cNvPr": [{ $: { id: "1", name: "" } }],
                        "p:cNvGrpSpPr": [{}],
                        "p:nvPr": [{}],
                    }],
                    "p:grpSpPr": [{
                        "a:xfrm": [{
                            "a:off": [{ $: { x: "0", y: "0" } }],
                            "a:ext": [{ $: { cx: "0", cy: "0" } }],
                            "a:chOff": [{ $: { x: "0", y: "0" } }],
                            "a:chExt": [{ $: { cx: "0", cy: "0" } }],
                        }],
                    }],
                    "p:sp": [
                        {
                            "p:nvSpPr": [{
                                "p:cNvPr": [{ $: { id: "2", name: `Slide Image Placeholder ${slideNumber}` } }],
                                "p:cNvSpPr": [{ "a:spLocks": [{ $: { noGrp: "1", noRot: "1", noChangeAspect: "1" } }] }],
                                "p:nvPr": [{ "p:ph": [{ $: { type: "sldImg" } }] }],
                            }],
                            "p:spPr": [{}],
                        },
                        notesBodyShape(paragraphs),
                    ],
                }],
            }],
            "p:clrMapOvr": [{ "a:masterClrMapping": [{}] }],
        },
    };
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

async function removeNotesMasterIdList(zip) {
    const presentationXml = await parseZipXml(zip, "ppt/presentation.xml");

    if (!presentationXml?.["p:presentation"]?.["p:notesMasterIdLst"]) {
        return;
    }

    delete presentationXml["p:presentation"]["p:notesMasterIdLst"];
    writeZipXml(zip, "ppt/presentation.xml", presentationXml);
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

// Parses the text format written by exportSlideNotesText. Splits on
// "Slide #N:" header lines, so blank lines inside notes are preserved.
// Known limitation: a notes line that itself starts with "Slide #N:" is
// indistinguishable from a header.
export function parseSpeakerNotesText(text) {
    const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    const headers = [...normalized.matchAll(/^Slide #(\d+):.*$/gm)];

    if (headers.length === 0) {
        throw new Error("Notes file has no \"Slide #N:\" headers; expected a file created by export-notes");
    }

    const preamble = normalized.slice(0, headers[0].index);

    if (preamble.trim() !== "") {
        throw new Error("Notes file has content before the first \"Slide #N:\" header");
    }

    const notesBySlideNumber = new Map();

    for (let index = 0; index < headers.length; index++) {
        const header = headers[index];
        const slideNumber = Number(header[1]);
        const blockStart = header.index + header[0].length;
        const blockEnd = index + 1 < headers.length ? headers[index + 1].index : normalized.length;

        if (notesBySlideNumber.has(slideNumber)) {
            throw new Error(`Notes file lists Slide #${slideNumber} more than once`);
        }

        notesBySlideNumber.set(slideNumber, normalized.slice(blockStart, blockEnd).trim());
    }

    return notesBySlideNumber;
}

async function ensureNotesMaster(zip) {
    const presentationRelsFile = "ppt/_rels/presentation.xml.rels";
    const presentationRels = await parseZipXml(zip, presentationRelsFile);
    const existingRelationship = findRelationshipByType(presentationRels, NOTES_MASTER_REL_TYPE);

    if (existingRelationship) {
        const existingPath = relationshipTargetToZipPath(presentationRelsFile, existingRelationship.$.Target);

        if (zip.files[existingPath]) {
            return existingPath;
        }
    }

    const presentationXml = await parseZipXml(zip, "ppt/presentation.xml");

    if (!presentationXml?.["p:presentation"] || !presentationRels?.Relationships) {
        throw new Error("Cannot create a notes master: presentation.xml or its relationships are missing");
    }

    const themeRelationship = findRelationshipByType(presentationRels, THEME_REL_TYPE);
    const sourceThemePath = themeRelationship
        ? relationshipTargetToZipPath(presentationRelsFile, themeRelationship.$.Target)
        : Object.keys(zip.files).find((fileName) => /^ppt\/theme\/theme\d+\.xml$/.test(fileName));

    if (!sourceThemePath || !zip.files[sourceThemePath]) {
        throw new Error("Cannot create a notes master: no theme part found in the presentation");
    }

    // The notes master gets its own copy of the theme rather than sharing the
    // slide master's part. PowerPoint offers to "repair" decks whose notes
    // master shares a theme when notesMasterIdLst sits in its schema position,
    // and genuine PowerPoint files always give each master a separate theme.
    const themeNumber = nextPartNumber(zip, /^ppt\/theme\/theme(\d+)\.xml$/);
    const themePath = `ppt/theme/theme${themeNumber}.xml`;

    zip.file(themePath, await zip.files[sourceThemePath].async("string"));
    await addContentTypeOverride(zip, `/${themePath}`, THEME_CONTENT_TYPE);

    const masterNumber = nextPartNumber(zip, /^ppt\/notesMasters\/notesMaster(\d+)\.xml$/);
    const masterPath = `ppt/notesMasters/notesMaster${masterNumber}.xml`;
    const masterRels = { Relationships: { $: { xmlns: RELATIONSHIPS_XMLNS } } };

    zip.file(masterPath, NOTES_MASTER_XML);
    appendRelationship(masterRels, {
        id: "rId1",
        type: THEME_REL_TYPE,
        target: path.posix.relative("ppt/notesMasters", themePath),
    });
    writeZipXml(zip, `ppt/notesMasters/_rels/notesMaster${masterNumber}.xml.rels`, masterRels);
    await addContentTypeOverride(zip, `/${masterPath}`, NOTES_MASTER_CONTENT_TYPE);

    const relationshipId = nextRelationshipId(presentationRels);

    appendRelationship(presentationRels, {
        id: relationshipId,
        type: NOTES_MASTER_REL_TYPE,
        target: `notesMasters/notesMaster${masterNumber}.xml`,
    });
    writeZipXml(zip, presentationRelsFile, presentationRels);

    const presentationElement = presentationXml["p:presentation"];
    const notesMasterId = { "p:notesMasterId": [{ $: { "r:id": relationshipId } }] };

    if (presentationElement["p:notesMasterIdLst"]) {
        const idList = presentationElement["p:notesMasterIdLst"][0];

        idList["p:notesMasterId"] = [
            ...asArray(idList["p:notesMasterId"]),
            notesMasterId["p:notesMasterId"][0],
        ];
    } else {
        presentationXml["p:presentation"] = insertPresentationChildInSchemaOrder(
            presentationElement,
            "p:notesMasterIdLst",
            [notesMasterId],
        );
    }

    writeZipXml(zip, "ppt/presentation.xml", presentationXml);

    return masterPath;
}

async function createNotesSlideForSlide(zip, slideFile, paragraphs, notesMasterPath) {
    const notesNumber = nextPartNumber(zip, /^ppt\/notesSlides\/notesSlide(\d+)\.xml$/);
    const notesPath = `ppt/notesSlides/notesSlide${notesNumber}.xml`;
    const notesRels = { Relationships: { $: { xmlns: RELATIONSHIPS_XMLNS } } };
    const slideNumber = slideNumberFromPath(slideFile);

    writeZipXml(zip, notesPath, buildNotesSlideXml(paragraphs, slideNumber));
    appendRelationship(notesRels, {
        id: "rId1",
        type: NOTES_MASTER_REL_TYPE,
        target: path.posix.relative("ppt/notesSlides", notesMasterPath),
    });
    appendRelationship(notesRels, {
        id: "rId2",
        type: SLIDE_REL_TYPE,
        target: `../slides/${path.posix.basename(slideFile)}`,
    });
    writeZipXml(zip, `ppt/notesSlides/_rels/notesSlide${notesNumber}.xml.rels`, notesRels);
    await addContentTypeOverride(zip, `/${notesPath}`, NOTES_SLIDE_CONTENT_TYPE);

    const slideRelsFile = slideRelationshipPath(slideFile);
    const slideRels = await ensureRelsXml(zip, slideRelsFile);

    appendRelationship(slideRels, {
        id: nextRelationshipId(slideRels),
        type: NOTES_SLIDE_REL_TYPE,
        target: `../notesSlides/notesSlide${notesNumber}.xml`,
    });
    writeZipXml(zip, slideRelsFile, slideRels);
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
        const slideFiles = sortedSlideFiles(zip);

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

        await removeNotesMasterIdList(zip);
        await removeNotesContentTypes(zip);
        await writeZipToFile(zip, outputPath);

        return outputPath;
    }

    async importSlideNotesText(pptxPath, notesTextPath, outputPath, options = {}) {
        const { onProgress } = options;
        let notesText;

        try {
            notesText = await readFile(notesTextPath, "utf8");
        } catch (error) {
            if (error.code === "ENOENT") {
                throw new Error(`Notes file not found: ${notesTextPath}`);
            }

            throw error;
        }

        const notesBySlideNumber = parseSpeakerNotesText(notesText);

        reportProgress(onProgress, { phase: "read-file", current: 0, total: 1 });

        const zip = await JSZip.loadAsync(await readFile(pptxPath));

        reportProgress(onProgress, { phase: "read-file", current: 1, total: 1 });

        const slideFiles = sortedSlideFiles(zip);
        const unknownSlides = [...notesBySlideNumber.keys()]
            .filter((slideNumber) => slideNumber < 1 || slideNumber > slideFiles.length);

        if (unknownSlides.length > 0) {
            const labels = unknownSlides.map((slideNumber) => `Slide #${slideNumber}`).join(", ");

            throw new Error(`${labels} not found in presentation (${slideFiles.length} slides)`);
        }

        const entries = [...notesBySlideNumber.entries()].sort(([left], [right]) => left - right);
        let notesMasterPath = null;

        reportProgress(onProgress, { phase: "import-notes", current: 0, total: entries.length });

        for (let index = 0; index < entries.length; index++) {
            const [slideNumber, text] = entries[index];
            const slideFile = slideFiles[slideNumber - 1];
            const notesFile = await this.findNotesFileForSlide(zip, slideFile);

            if (notesFile) {
                const notesXml = await parseZipXml(zip, notesFile);

                setNotesBodyParagraphs(notesXml, notesParagraphsFromText(text));
                writeZipXml(zip, notesFile, notesXml);
            } else if (text !== "") {
                notesMasterPath = notesMasterPath || await ensureNotesMaster(zip);
                await createNotesSlideForSlide(zip, slideFile, notesParagraphsFromText(text), notesMasterPath);
            }

            reportProgress(onProgress, { phase: "import-notes", current: index + 1, total: entries.length });
        }

        await writeZipToFile(zip, outputPath);

        return outputPath;
    }
}
