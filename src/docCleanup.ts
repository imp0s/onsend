import JSZip from 'jszip';
import { XMLBuilder, XMLParser } from 'fast-xml-parser';

export interface CleanupSummary {
  updated: boolean;
  removedComments: boolean;
  removedMetadata: boolean;
}

const xmlOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: '',
  suppressBooleanAttributes: false
};

function removeCoreProperties(parser: XMLParser, builder: XMLBuilder, zip: JSZip): Promise<boolean> {
  const corePath = 'docProps/core.xml';
  const file = zip.file(corePath);
  if (!file) return Promise.resolve(false);

  return file.async('string').then((xml) => {
    const parsed = parser.parse(xml);
    const props = parsed['cp:coreProperties'] ?? parsed.coreProperties;

    if (!props) return false;
    const targets = ['dc:creator', 'cp:lastModifiedBy', 'cp:lastPrinted', 'dcterms:created', 'dcterms:modified'];
    let changed = false;

    targets.forEach((prop) => {
      if (props[prop]) {
        delete props[prop];
        changed = true;
      }
    });

    if (changed) {
      const rootKey = parsed['cp:coreProperties'] ? 'cp:coreProperties' : 'coreProperties';
      const updated = builder.build({ [rootKey]: props });
      zip.file(corePath, updated);
    }

    return changed;
  });
}

function removeComments(parser: XMLParser, builder: XMLBuilder, zip: JSZip): Promise<boolean> {
  const commentsPath = 'word/comments.xml';
  const file = zip.file(commentsPath);
  if (!file) return Promise.resolve(false);

  return file.async('string').then((xml) => {
    const parsed = parser.parse(xml);
    const commentsRoot = parsed['w:comments'] ?? parsed.comments;
    if (!commentsRoot || !commentsRoot['w:comment'] && !commentsRoot.comment) return false;

    if (Array.isArray(commentsRoot['w:comment'])) {
      commentsRoot['w:comment'] = [];
    }
    if (Array.isArray(commentsRoot.comment)) {
      commentsRoot.comment = [];
    }

    const rootKey = parsed['w:comments'] ? 'w:comments' : 'comments';
    const updated = builder.build({ [rootKey]: commentsRoot });
    zip.file(commentsPath, updated);
    return true;
  });
}

export async function removeMetadataAndComments(buffer: ArrayBuffer | Uint8Array | Buffer): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(buffer);
  const parser = new XMLParser(xmlOptions);
  const builder = new XMLBuilder({ ...xmlOptions, format: true });

  const [metadataRemoved, commentsRemoved] = await Promise.all([
    removeCoreProperties(parser, builder, zip),
    removeComments(parser, builder, zip)
  ]);

  if (!metadataRemoved && !commentsRemoved) {
    return new Uint8Array(buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer));
  }

  return zip.generateAsync({ type: 'uint8array' });
}

export function base64ToUint8Array(base64: string): Uint8Array {
  const raw = Buffer.from(base64, 'base64');
  return new Uint8Array(raw.buffer, raw.byteOffset, raw.byteLength);
}

export function uint8ArrayToBase64(data: Uint8Array): string {
  return Buffer.from(data).toString('base64');
}
