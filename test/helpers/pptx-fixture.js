import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import JSZip from "jszip";

function escapeXml(value) {
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll("\"", "&quot;")
        .replaceAll("'", "&apos;");
}

function textRunsXml(textRuns) {
    return textRuns
        .map((text) => `<a:r><a:t>${escapeXml(text)}</a:t></a:r>`)
        .join("");
}

function paragraphXml(textRuns) {
    return `<a:p>${textRunsXml(textRuns)}</a:p>`;
}

function relationshipXml(relationships) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${relationships.map((relationship) => `  <Relationship Id="${escapeXml(relationship.id)}" Type="${escapeXml(relationship.type)}" Target="${escapeXml(relationship.target)}"/>`).join("\n")}
</Relationships>`;
}

function contentTypesXml(slides) {
    const notesOverrides = slides
        .filter((slide) => slide.notes)
        .map((slide) => `  <Override PartName="/ppt/notesSlides/notesSlide${slide.number}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`)
        .join("\n");

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
${slides.map((slide) => `  <Override PartName="/ppt/slides/slide${slide.number}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("\n")}
${notesOverrides}
</Types>`;
}

export function slideXml({
    placeholderType = "title",
    textRuns = ["Test title"],
    includePlaceholder = true,
    bodyPlaceholderType = null,
} = {}) {
    const placeholder = includePlaceholder
        ? `<p:ph type="${escapeXml(placeholderType)}"/>`
        : "";
    const bodyPlaceholder = bodyPlaceholderType
        ? `<p:ph type="${escapeXml(bodyPlaceholderType)}"/>`
        : "";

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr/>
          <p:nvPr>${placeholder}</p:nvPr>
        </p:nvSpPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          ${paragraphXml(textRuns)}
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Body 1"/>
          <p:cNvSpPr/>
          <p:nvPr>${bodyPlaceholder}</p:nvPr>
        </p:nvSpPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          ${paragraphXml(["Body text"])}
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

export function notesSlideXml({
    notes = [["Speaker notes"]],
} = {}) {
    const paragraphs = notes.map(paragraphXml).join("\n          ");

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Notes Placeholder 2"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="body"/></p:nvPr>
        </p:nvSpPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          ${paragraphs}
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>`;
}

export async function writePptxFixture(slides) {
    const tempDir = await mkdtemp(path.join(tmpdir(), "slidenotes-"));
    const filePath = path.join(tempDir, "fixture.pptx");
    const zip = new JSZip();

    zip.file("[Content_Types].xml", contentTypesXml(slides));

    for (const slide of slides) {
        zip.file(`ppt/slides/slide${slide.number}.xml`, slideXml(slide));

        if (slide.notes) {
            zip.file(`ppt/notesSlides/notesSlide${slide.number}.xml`, notesSlideXml(slide));
            zip.file(`ppt/slides/_rels/slide${slide.number}.xml.rels`, relationshipXml([
                {
                    id: "rId1",
                    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
                    target: `../notesSlides/notesSlide${slide.number}.xml`,
                },
            ]));
        }
    }

    await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));

    return filePath;
}
