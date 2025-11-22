const JSZip = require("jszip");
const { XMLBuilder, XMLParser } = require("fast-xml-parser");

const xmlOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "",
  suppressBooleanAttributes: false,
};

function removeCoreProperties(parser, builder, zip) {
  const corePath = "docProps/core.xml";
  const file = zip.file(corePath);
  if (!file) return Promise.resolve(false);

  return file.async("string").then((xml) => {
    const parsed = parser.parse(xml);
    const props = parsed["cp:coreProperties"] ?? parsed.coreProperties;

    if (!props) return false;
    const targets = [
      "dc:creator",
      "cp:lastModifiedBy",
      "cp:lastPrinted",
      "dcterms:created",
      "dcterms:modified",
    ];
    let changed = false;

    targets.forEach((prop) => {
      if (props[prop]) {
        delete props[prop];
        changed = true;
      }
    });

    if (changed) {
      const rootKey = parsed["cp:coreProperties"]
        ? "cp:coreProperties"
        : "coreProperties";
      const updated = builder.build({ [rootKey]: props });
      zip.file(corePath, updated);
    }

    return changed;
  });
}

function removeComments(parser, builder, zip) {
  const commentsPath = "word/comments.xml";
  const file = zip.file(commentsPath);
  if (!file) return Promise.resolve(false);

  return file.async("string").then((xml) => {
    const parsed = parser.parse(xml);
    const commentsRoot = parsed["w:comments"] ?? parsed.comments;
    if (!commentsRoot) return false;

    let changed = false;

    if (commentsRoot["w:comment"]) {
      delete commentsRoot["w:comment"];
      changed = true;
    }
    if (commentsRoot.comment) {
      delete commentsRoot.comment;
      changed = true;
    }

    if (!changed) return false;

    const rootKey = parsed["w:comments"] ? "w:comments" : "comments";
    const updated = builder.build({ [rootKey]: commentsRoot });
    zip.file(commentsPath, updated);
    return true;
  });
}

async function removeMetadataAndComments(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const parser = new XMLParser(xmlOptions);
  const builder = new XMLBuilder({ ...xmlOptions, format: true });

  const [metadataRemoved, commentsRemoved] = await Promise.all([
    removeCoreProperties(parser, builder, zip),
    removeComments(parser, builder, zip),
  ]);

  if (!metadataRemoved && !commentsRemoved) {
    return new Uint8Array(
      buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer),
    );
  }

  return zip.generateAsync({ type: "uint8array" });
}

function base64ToUint8Array(base64) {
  if (typeof Buffer !== "undefined") {
    const raw = Buffer.from(base64, "base64");
    return new Uint8Array(raw.buffer, raw.byteOffset, raw.byteLength);
  }

  if (typeof atob === "function") {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }

  throw new Error("No base64 decoder available");
}

function uint8ArrayToBase64(data) {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(data).toString("base64");
  }

  if (typeof btoa === "function") {
    let binary = "";
    for (let i = 0; i < data.byteLength; i += 1) {
      binary += String.fromCharCode(data[i]);
    }
    return btoa(binary);
  }

  throw new Error("No base64 encoder available");
}

module.exports = {
  removeMetadataAndComments,
  base64ToUint8Array,
  uint8ArrayToBase64,
};
