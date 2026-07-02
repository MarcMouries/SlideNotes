import { describe, expect, test } from "bun:test";
import { parseStringPromise as parseXml } from "xml2js";
import SlideNotes, { NO_TITLE } from "../src/index.js";
import { slideXml, writePptxFixture } from "./helpers/pptx-fixture.js";

describe("SlideNotes", () => {
    test("loads slides in numeric order and exposes titles", async () => {
        const filePath = await writePptxFixture([
            { number: 10, textRuns: ["Tenth"] },
            { number: 2, textRuns: ["Second"] },
            { number: 1, textRuns: ["First"] },
        ]);

        const reader = await new SlideNotes().load(filePath);

        expect(reader.slides).toHaveLength(3);
        expect(reader.slides.map((slide) => slide.getTitle())).toEqual([
            "First",
            "Second",
            "Tenth",
        ]);
    });

    test("resets slides when the same reader loads another presentation", async () => {
        const firstDeck = await writePptxFixture([
            { number: 1, textRuns: ["First deck"] },
            { number: 2, textRuns: ["Another slide"] },
        ]);
        const secondDeck = await writePptxFixture([
            { number: 1, textRuns: ["Second deck"] },
        ]);
        const reader = new SlideNotes();

        await reader.load(firstDeck);
        await reader.load(secondDeck);

        expect(reader.slides).toHaveLength(1);
        expect(reader.slides[0].getTitle()).toBe("Second deck");
    });

    test("reports progress while loading slides", async () => {
        const filePath = await writePptxFixture([
            { number: 1, textRuns: ["First"] },
            { number: 2, textRuns: ["Second"] },
        ]);
        const progressEvents = [];

        await new SlideNotes().load(filePath, {
            onProgress: (progress) => progressEvents.push(progress),
        });

        expect(progressEvents).toContainEqual({
            phase: "read-file",
            current: 1,
            total: 1,
        });
        expect(progressEvents).toContainEqual({
            phase: "read-slides",
            current: 2,
            total: 2,
        });
    });

    test("extracts title and centered title placeholders", async () => {
        const reader = new SlideNotes();
        const titleSlide = await parseXml(slideXml({ placeholderType: "title", textRuns: ["Title"] }));
        const centeredTitleSlide = await parseXml(slideXml({
            placeholderType: "ctrTitle",
            textRuns: ["Centered title"],
        }));

        expect(reader.extractTitle(titleSlide)).toBe("Title");
        expect(reader.extractTitle(centeredTitleSlide)).toBe("Centered title");
    });

    test("joins text split across XML runs", async () => {
        const reader = new SlideNotes();
        const parsedSlide = await parseXml(slideXml({
            textRuns: ["Annual ", "Review"],
        }));

        expect(reader.extractTitle(parsedSlide)).toBe("Annual Review");
    });

    test("uses the first shape text when no title placeholder exists", async () => {
        const reader = new SlideNotes();
        const parsedSlide = await parseXml(slideXml({
            includePlaceholder: false,
            textRuns: ["Body text"],
        }));

        expect(reader.extractTitle(parsedSlide)).toBe("Body text");
    });

    test("returns no title when the slide only has body placeholder text", async () => {
        const reader = new SlideNotes();
        const parsedSlide = await parseXml(slideXml({
            includePlaceholder: false,
            textRuns: [""],
            bodyPlaceholderType: "body",
        }));

        expect(reader.extractTitle(parsedSlide)).toBe(NO_TITLE);
    });
});
