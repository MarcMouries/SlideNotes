import path from "node:path";
import xml2js from "xml2js";

const { parseStringPromise: parseXml } = xml2js;

export const ZIP_METHOD_STORED = 0;
export const ZIP_METHOD_DEFLATED = 8;

const ZIP_END_OF_CENTRAL_DIR_SIGNATURE = Buffer.from([0x50, 0x4b, 0x05, 0x06]);

// Reads each central directory entry's compression method straight from the
// zip bytes, since JSZip does not expose them after loading.
export function zipCompressionMethods(buffer) {
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

function asArray(value) {
    if (value === undefined || value === null) {
        return [];
    }

    return Array.isArray(value) ? value : [value];
}

async function relationshipsIn(zip, relsFile) {
    const relsXml = await parseXml(await zip.files[relsFile].async("string"));

    return asArray(relsXml.Relationships.Relationship);
}

// Returns [{relsFile, target}] for internal relationship targets that do not
// resolve to a part in the package. A valid package returns [].
export async function unresolvedRelationshipTargets(zip) {
    const relsFiles = Object.keys(zip.files).filter((fileName) => fileName.endsWith(".rels"));
    const unresolved = [];

    for (const relsFile of relsFiles) {
        for (const relationship of await relationshipsIn(zip, relsFile)) {
            if (relationship.$.TargetMode === "External") {
                continue;
            }

            const sourceDir = path.posix.dirname(path.posix.dirname(relsFile));
            const target = relationship.$.Target;
            const targetPath = target.startsWith("/")
                ? target.slice(1)
                : path.posix.normalize(path.posix.join(sourceDir, target));

            if (!zip.files[targetPath]) {
                unresolved.push({ relsFile, target });
            }
        }
    }

    return unresolved;
}

// Returns r:id / r:embed values referenced by a part's XML that have no
// matching relationship Id in its .rels file. A valid part returns [].
export async function danglingRelationshipIds(zip, partFile) {
    const partDir = path.posix.dirname(partFile);
    const relsFile = `${partDir}/_rels/${path.posix.basename(partFile)}.rels`;
    const partXml = await zip.files[partFile].async("string");
    const referencedIds = [...partXml.matchAll(/r:(?:id|embed)="([^"]+)"/g)].map((match) => match[1]);
    const relationshipIds = zip.files[relsFile]
        ? new Set((await relationshipsIn(zip, relsFile)).map((relationship) => relationship.$.Id))
        : new Set();

    return referencedIds.filter((id) => !relationshipIds.has(id));
}

// Returns [{relsFile, id}] for relationship Ids declared more than once in
// the same .rels file. A valid package returns [].
export async function duplicateRelationshipIds(zip) {
    const relsFiles = Object.keys(zip.files).filter((fileName) => fileName.endsWith(".rels"));
    const duplicates = [];

    for (const relsFile of relsFiles) {
        const seen = new Set();

        for (const relationship of await relationshipsIn(zip, relsFile)) {
            if (seen.has(relationship.$.Id)) {
                duplicates.push({ relsFile, id: relationship.$.Id });
            }

            seen.add(relationship.$.Id);
        }
    }

    return duplicates;
}
