import { ZipEntry } from './types';

const textEncoder = new TextEncoder();
const CRC32_TABLE = createCrc32Table();

/**
 * Precomputes the CRC32 table used for ZIP checksums.
 */
function createCrc32Table(): Uint32Array {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) {
      c = (c & 1) !== 0 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
}

/**
 * Computes CRC32 checksum for a byte buffer.
 */
function crc32(buffer: Uint8Array): number {
  let crc = 0xffffffff;
  for (let i = 0; i < buffer.length; i += 1) {
    crc = CRC32_TABLE[(crc ^ buffer[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Encodes a string into UTF-8 bytes.
 */
export function encodeUtf8(value: string): Uint8Array {
  return textEncoder.encode(value);
}

/**
 * Writes a 32-bit unsigned integer into a byte buffer in little-endian order.
 */
function writeUint32(buffer: Uint8Array, offset: number, value: number): void {
  buffer[offset] = value & 0xff;
  buffer[offset + 1] = (value >>> 8) & 0xff;
  buffer[offset + 2] = (value >>> 16) & 0xff;
  buffer[offset + 3] = (value >>> 24) & 0xff;
}

/**
 * Builds a minimal ZIP archive from the given entries.
 */
export function buildZip(entries: ZipEntry[]): Uint8Array {
  const encoder = new TextEncoder();
  const centralDirectoryEntries: Uint8Array[] = [];
  const localFileRecords: Uint8Array[] = [];
  let offset = 0;

  entries.forEach((entry) => {
    const nameBytes = encoder.encode(entry.name);
    const data = entry.data;
    const crc = crc32(data);
    const compressedSize = data.length;
    const uncompressedSize = data.length;

    const localHeader = new Uint8Array(30 + nameBytes.length);
    writeUint32(localHeader, 0, 0x04034b50);
    writeUint32(localHeader, 4, 20);
    writeUint32(localHeader, 6, 0);
    writeUint32(localHeader, 8, 0);
    writeUint32(localHeader, 10, 0);
    writeUint32(localHeader, 14, crc);
    writeUint32(localHeader, 18, compressedSize);
    writeUint32(localHeader, 22, uncompressedSize);
    writeUint32(localHeader, 26, nameBytes.length);
    writeUint32(localHeader, 28, 0);
    localHeader.set(nameBytes, 30);

    const localFileRecord = new Uint8Array(localHeader.length + data.length);
    localFileRecord.set(localHeader, 0);
    localFileRecord.set(data, localHeader.length);
    localFileRecords.push(localFileRecord);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    writeUint32(centralHeader, 0, 0x02014b50);
    writeUint32(centralHeader, 4, 0);
    writeUint32(centralHeader, 6, 20);
    writeUint32(centralHeader, 8, 20);
    writeUint32(centralHeader, 10, 0);
    writeUint32(centralHeader, 12, 0);
    writeUint32(centralHeader, 14, 0);
    writeUint32(centralHeader, 18, crc);
    writeUint32(centralHeader, 22, compressedSize);
    writeUint32(centralHeader, 26, uncompressedSize);
    writeUint32(centralHeader, 30, nameBytes.length);
    writeUint32(centralHeader, 32, 0);
    writeUint32(centralHeader, 34, 0);
    writeUint32(centralHeader, 36, 0);
    writeUint32(centralHeader, 38, 0);
    writeUint32(centralHeader, 42, 0);
    writeUint32(centralHeader, 44, offset);
    centralHeader.set(nameBytes, 46);

    centralDirectoryEntries.push(centralHeader);
    offset += localFileRecord.length;
  });

  const centralSize = centralDirectoryEntries.reduce((sum, item) => sum + item.length, 0);
  const totalSize = offset + centralSize + 22;
  const output = new Uint8Array(totalSize);

  let cursor = 0;
  localFileRecords.forEach((record) => {
    output.set(record, cursor);
    cursor += record.length;
  });

  const startOfCentral = cursor;
  centralDirectoryEntries.forEach((record) => {
    output.set(record, cursor);
    cursor += record.length;
  });

  writeUint32(output, cursor, 0x06054b50);
  writeUint32(output, cursor + 4, 0);
  writeUint32(output, cursor + 6, 0);
  writeUint32(output, cursor + 8, entries.length);
  writeUint32(output, cursor + 12, entries.length);
  writeUint32(output, cursor + 16, centralSize);
  writeUint32(output, cursor + 20, startOfCentral);

  return output;
}
