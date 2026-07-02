#!/usr/bin/env bun

import path from "node:path";
import { writeDemoPresentationWithNotes } from "../test/helpers/demo-presentation.js";

const fixturePath = path.resolve("test/fixture/simple-with-notes.pptx");

await writeDemoPresentationWithNotes(fixturePath);

console.log(`Generated ${fixturePath}`);
console.log("Try:");
console.log("  slidenotes list ./test/fixture/simple-with-notes.pptx");
console.log("  slidenotes export-notes ./test/fixture/simple-with-notes.pptx");
console.log("  slidenotes remove-notes ./test/fixture/simple-with-notes.pptx");
