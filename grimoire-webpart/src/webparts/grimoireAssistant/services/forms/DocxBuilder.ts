interface IZipEntry {
  name: string;
  data: Uint8Array;
}

const CRC32_TABLE: Uint32Array = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let value = i;
    for (let j = 0; j < 8; j++) {
      value = (value & 1) === 1 ? (0xedb88320 ^ (value >>> 1)) : (value >>> 1);
    }
    table[i] = value >>> 0;
  }
  return table;
})();

function encodeUtf8(value: string): Uint8Array {
  if (typeof TextEncoder !== 'undefined') {
    return new TextEncoder().encode(value);
  }

  const bufferCtor = (globalThis as { Buffer?: { from: (data: string, encoding: string) => Uint8Array } }).Buffer;
  if (bufferCtor) {
    return bufferCtor.from(value, 'utf8');
  }

  const encoded = unescape(encodeURIComponent(value));
  const bytes = new Uint8Array(encoded.length);
  for (let i = 0; i < encoded.length; i++) {
    bytes[i] = encoded.charCodeAt(i);
  }
  return bytes;
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function toDosTime(date: Date): number {
  return ((date.getHours() & 0x1f) << 11)
    | ((date.getMinutes() & 0x3f) << 5)
    | Math.floor(date.getSeconds() / 2);
}

function toDosDate(date: Date): number {
  const year = Math.max(1980, date.getFullYear());
  return (((year - 1980) & 0x7f) << 9)
    | (((date.getMonth() + 1) & 0x0f) << 5)
    | (date.getDate() & 0x1f);
}

function writeUint16(target: Uint8Array, offset: number, value: number): void {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >>> 8) & 0xff;
}

function writeUint32(target: Uint8Array, offset: number, value: number): void {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >>> 8) & 0xff;
  target[offset + 2] = (value >>> 16) & 0xff;
  target[offset + 3] = (value >>> 24) & 0xff;
}

function concatUint8Arrays(parts: Uint8Array[]): Uint8Array {
  const totalLength = parts.reduce((sum, part) => sum + part.length, 0);
  const merged = new Uint8Array(totalLength);
  let offset = 0;
  parts.forEach((part) => {
    merged.set(part, offset);
    offset += part.length;
  });
  return merged;
}

function crc32(bytes: Uint8Array): number {
  let value = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) {
    value = CRC32_TABLE[(value ^ bytes[i]) & 0xff] ^ (value >>> 8);
  }
  return (value ^ 0xffffffff) >>> 0;
}

function buildStoredZip(entries: IZipEntry[], modifiedAt: Date): Uint8Array {
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  const dosTime = toDosTime(modifiedAt);
  const dosDate = toDosDate(modifiedAt);
  let localOffset = 0;

  entries.forEach((entry) => {
    const fileName = encodeUtf8(entry.name);
    const fileData = entry.data;
    const checksum = crc32(fileData);

    const localHeader = new Uint8Array(30 + fileName.length);
    writeUint32(localHeader, 0, 0x04034b50);
    writeUint16(localHeader, 4, 20);
    writeUint16(localHeader, 6, 0);
    writeUint16(localHeader, 8, 0);
    writeUint16(localHeader, 10, dosTime);
    writeUint16(localHeader, 12, dosDate);
    writeUint32(localHeader, 14, checksum);
    writeUint32(localHeader, 18, fileData.length);
    writeUint32(localHeader, 22, fileData.length);
    writeUint16(localHeader, 26, fileName.length);
    writeUint16(localHeader, 28, 0);
    localHeader.set(fileName, 30);

    localParts.push(localHeader, fileData);

    const centralHeader = new Uint8Array(46 + fileName.length);
    writeUint32(centralHeader, 0, 0x02014b50);
    writeUint16(centralHeader, 4, 20);
    writeUint16(centralHeader, 6, 20);
    writeUint16(centralHeader, 8, 0);
    writeUint16(centralHeader, 10, 0);
    writeUint16(centralHeader, 12, dosTime);
    writeUint16(centralHeader, 14, dosDate);
    writeUint32(centralHeader, 16, checksum);
    writeUint32(centralHeader, 20, fileData.length);
    writeUint32(centralHeader, 24, fileData.length);
    writeUint16(centralHeader, 28, fileName.length);
    writeUint16(centralHeader, 30, 0);
    writeUint16(centralHeader, 32, 0);
    writeUint16(centralHeader, 34, 0);
    writeUint16(centralHeader, 36, 0);
    writeUint32(centralHeader, 38, 0);
    writeUint32(centralHeader, 42, localOffset);
    centralHeader.set(fileName, 46);

    centralParts.push(centralHeader);
    localOffset += localHeader.length + fileData.length;
  });

  const centralDirectory = concatUint8Arrays(centralParts);
  const endRecord = new Uint8Array(22);
  writeUint32(endRecord, 0, 0x06054b50);
  writeUint16(endRecord, 4, 0);
  writeUint16(endRecord, 6, 0);
  writeUint16(endRecord, 8, entries.length);
  writeUint16(endRecord, 10, entries.length);
  writeUint32(endRecord, 12, centralDirectory.length);
  writeUint32(endRecord, 16, localOffset);
  writeUint16(endRecord, 20, 0);

  return concatUint8Arrays([...localParts, centralDirectory, endRecord]);
}

function buildParagraphsXml(content: string): string {
  const normalized = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = normalized.split('\n');
  if (lines.length === 0) {
    return '<w:p/>';
  }

  return lines
    .map((line) => {
      if (!line) {
        return '<w:p/>';
      }
      return `<w:p><w:r><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`;
    })
    .join('');
}

function encodeBase64(bytes: Uint8Array): string {
  const bufferCtor = (globalThis as { Buffer?: { from: (data: Uint8Array) => { toString: (encoding: string) => string } } }).Buffer;
  if (bufferCtor) {
    return bufferCtor.from(bytes).toString('base64');
  }

  if (typeof btoa === 'function') {
    let binary = '';
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      binary += String.fromCharCode(...Array.from(chunk));
    }
    return btoa(binary);
  }

  throw new Error('Base64 encoding is not available in this environment.');
}

export function buildDocxBase64FromText(content: string): string {
  const modifiedAt = new Date();
  const isoTimestamp = modifiedAt.toISOString();
  const paragraphs = buildParagraphsXml(content);

  const files: IZipEntry[] = [
    {
      name: '[Content_Types].xml',
      data: encodeUtf8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/>'
        + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        + '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        + '</Types>'
      )
    },
    {
      name: '_rels/.rels',
      data: encodeUtf8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        + '</Relationships>'
      )
    },
    {
      name: 'docProps/core.xml',
      data: encodeUtf8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
        + ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
        + ' xmlns:dcterms="http://purl.org/dc/terms/"'
        + ' xmlns:dcmitype="http://purl.org/dc/dcmitype/"'
        + ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        + '<dc:title>Grimoire Document</dc:title>'
        + '<dc:creator>Grimoire</dc:creator>'
        + '<cp:lastModifiedBy>Grimoire</cp:lastModifiedBy>'
        + `<dcterms:created xsi:type="dcterms:W3CDTF">${isoTimestamp}</dcterms:created>`
        + `<dcterms:modified xsi:type="dcterms:W3CDTF">${isoTimestamp}</dcterms:modified>`
        + '</cp:coreProperties>'
      )
    },
    {
      name: 'docProps/app.xml',
      data: encodeUtf8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
        + ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        + '<Application>Grimoire</Application>'
        + '</Properties>'
      )
    },
    {
      name: 'word/document.xml',
      data: encodeUtf8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + '<w:body>'
        + paragraphs
        + '<w:sectPr>'
        + '<w:pgSz w:w="12240" w:h="15840"/>'
        + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
        + '</w:sectPr>'
        + '</w:body>'
        + '</w:document>'
      )
    }
  ];

  return encodeBase64(buildStoredZip(files, modifiedAt));
}
