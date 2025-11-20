import JSZip from "jszip";
import {
  base64ToUint8Array,
  removeMetadataAndComments,
  uint8ArrayToBase64,
} from "../src/docCleanup";

async function buildDocx(
  withComments = true,
  withMetadata = true,
): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
  );
  if (withMetadata) {
    zip.file(
      "docProps/core.xml",
      `<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:creator>Unit Tester</dc:creator>
  <cp:lastModifiedBy>Another User</cp:lastModifiedBy>
  <dcterms:created>2024-01-01T00:00:00Z</dcterms:created>
</cp:coreProperties>`,
    );
  }

  if (withComments) {
    zip.file(
      "word/comments.xml",
      `<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Unit Tester">
    <w:p><w:r><w:t>Test comment</w:t></w:r></w:p>
  </w:comment>
</w:comments>`,
    );
  }

  return zip.generateAsync({ type: "uint8array" });
}

describe("removeMetadataAndComments", () => {
  it("strips metadata and comments from docx files", async () => {
    const sample = await buildDocx(true, true);
    const cleaned = await removeMetadataAndComments(sample);
    const zip = await JSZip.loadAsync(cleaned);

    const coreXml = await zip.file("docProps/core.xml")!.async("string");
    expect(coreXml).not.toContain("dc:creator");
    expect(coreXml).not.toContain("cp:lastModifiedBy");
    expect(coreXml).not.toContain("dcterms:created");

    const commentsXml = await zip.file("word/comments.xml")!.async("string");
    expect(commentsXml).not.toMatch(/<w:comment\b/);
  });

  it("leaves files unchanged when nothing is removable", async () => {
    const sample = await buildDocx(false, false);
    const cleaned = await removeMetadataAndComments(sample);
    expect(cleaned).toEqual(sample);
  });
});

describe("base64 helpers", () => {
  it("converts to and from base64 accurately", () => {
    const data = Uint8Array.from([1, 2, 3, 4, 5]);
    const b64 = uint8ArrayToBase64(data);
    const roundTripped = base64ToUint8Array(b64);
    expect(Array.from(roundTripped)).toEqual(Array.from(data));
  });
});
