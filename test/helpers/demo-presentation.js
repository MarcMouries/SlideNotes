import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import pptxgen from "pptxgenjs";

export const demoSlides = [
    {
        title: "Opening",
        body: "A simple generated deck for testing speaker notes.",
        notes: "Welcome everyone.\nHere is the plan.",
    },
    {
        title: "Budget",
        body: "This slide should remain after notes are removed.",
        notes: "Mention Q4 numbers.\nPause for questions.",
    },
    {
        title: null,
        body: null,
        notes: "This note belongs to a slide without a title.",
        expectedTitle: "No title",
    },
    {
        title: "Appendix",
        body: "This slide intentionally has no speaker notes.",
        notes: "",
    },
];

// pptxgenjs ignores its `compression` option in Node, so rewrite the package
// deflated to match how PowerPoint itself saves files.
async function recompressPptx(pptxPath) {
    const zip = await JSZip.loadAsync(await readFile(pptxPath));

    await writeFile(pptxPath, await zip.generateAsync({
        type: "nodebuffer",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
    }));
}

export async function writeDemoPresentationWithNotes(outputPath) {
    await mkdir(path.dirname(outputPath), { recursive: true });

    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "SlideNotes tests";
    pptx.company = "SlideNotes";
    pptx.subject = "Speaker notes workflow";
    pptx.title = "SlideNotes Notes Demo";

    for (const demoSlide of demoSlides) {
        const slide = pptx.addSlide();

        slide.background = { color: "F7F9FC" };
        if (demoSlide.title) {
            slide.addText(demoSlide.title, {
                x: 0.6,
                y: 0.45,
                w: 11,
                h: 0.6,
                fontFace: "Aptos Display",
                fontSize: 32,
                bold: true,
                color: "1F2937",
            });
        }
        if (demoSlide.body) {
            slide.addText(demoSlide.body, {
                x: 0.75,
                y: 1.45,
                w: 10.8,
                h: 0.7,
                fontFace: "Aptos",
                fontSize: 18,
                color: "334155",
            });
            slide.addText("Speaker notes are attached to this slide.", {
                x: 0.75,
                y: 2.55,
                w: 8.5,
                h: 0.4,
                fontFace: "Aptos",
                fontSize: 14,
                italic: true,
                color: "64748B",
            });
        }
        if (demoSlide.notes) {
            slide.addNotes(demoSlide.notes);
        }
    }

    await pptx.writeFile({ fileName: outputPath });
    await recompressPptx(outputPath);

    return outputPath;
}
