#!/usr/bin/env bun

import path from "node:path";
import { exit } from "node:process";
import SlideNotes from "../src/index.js";

const PROGRESS_LABELS = {
    "read-file": "Reading file",
    "read-slides": "Reading slides",
    "export-notes": "Exporting notes",
};

class ProgressBar {
    constructor(stream = process.stderr) {
        this.stream = stream;
        this.enabled = Boolean(stream.isTTY);
        this.lastLineLength = 0;
    }

    update({ phase, current, total }) {
        if (!this.enabled || total < 1) {
            return;
        }

        const label = PROGRESS_LABELS[phase] || phase;
        const width = 28;
        const ratio = Math.min(current / total, 1);
        const filled = Math.round(ratio * width);
        const empty = width - filled;
        const percent = Math.round(ratio * 100);
        const line = `${label} [${"#".repeat(filled)}${"-".repeat(empty)}] ${current}/${total} ${percent}%`;

        this.stream.write(`\r${line}${" ".repeat(Math.max(this.lastLineLength - line.length, 0))}`);
        this.lastLineLength = line.length;
    }

    finish() {
        if (!this.enabled || this.lastLineLength === 0) {
            return;
        }

        this.stream.write("\n");
        this.lastLineLength = 0;
    }
}

function printUsage() {
    console.error("Usage:");
    console.error("  slidenotes list <pptx file path>");
    console.error("  slidenotes export-notes <pptx file path>");
    console.error("  slidenotes remove-notes <pptx file path>");
}

function outputPathFor(pptxPath, suffix, extension) {
    const parsedPath = path.parse(pptxPath);

    return path.join(parsedPath.dir, `${parsedPath.name}${suffix}${extension}`);
}

function notesOutputPathFor(pptxPath) {
    return outputPathFor(pptxPath, "-speaker-notes", ".txt");
}

function withoutNotesOutputPathFor(pptxPath) {
    return outputPathFor(pptxPath, "-without-notes", ".pptx");
}

const [command, ...args] = process.argv.slice(2);

if (!command) {
    printUsage();
    exit(1);
}

const reader = new SlideNotes();
const progressBar = new ProgressBar();

async function printSlides(filePath) {
    const pres = await reader.load(filePath, {
        onProgress: (progress) => progressBar.update(progress),
    });

    progressBar.finish();

    const slideCount = pres.slides.length;
    const slideCountDigits = String(slideCount).length;

    console.log(`Total slides: ${slideCount}\n`);

    for (let i = 0; i < slideCount; i++) {
        const slide = pres.slides[i];
        const title = slide.getTitle();
        const slideNumber = String(i + 1).padStart(slideCountDigits, " ");

        console.log(`${slideNumber}: ${title}`);
    }
}

try {
    if (command === "list" && args.length === 1) {
        await printSlides(args[0]);
    } else if (command === "export-notes" && args.length === 1) {
        const outputPath = notesOutputPathFor(args[0]);

        await reader.exportSlideNotesText(args[0], outputPath, {
            onProgress: (progress) => progressBar.update(progress),
        });

        progressBar.finish();
        console.log(`Wrote slide notes to ${outputPath}`);
    } else if (command === "remove-notes" && args.length === 1) {
        const outputPath = withoutNotesOutputPathFor(args[0]);

        await reader.removeNotes(args[0], outputPath);
        console.log(`Wrote presentation without notes to ${outputPath}`);
    } else {
        printUsage();
        exit(1);
    }
} catch (error) {
    progressBar.finish();
    console.error("Error reading PowerPoint file:", error);
    exit(1);
}
