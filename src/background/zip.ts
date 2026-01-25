type ZipFile = { path: string; data: Uint8Array };

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(bytes: Uint8Array) {
  let crc = 0 ^ -1;
  for (let i = 0; i < bytes.length; i++) {
    const b = bytes[i];
    crc = (crc >>> 8) ^ crcTable[(crc ^ b) & 0xff];
  }
  return (crc ^ -1) >>> 0;
}

function writeUint16(buf: Uint8Array, offset: number, value: number) {
  buf[offset] = value & 0xff;
  buf[offset + 1] = (value >>> 8) & 0xff;
}

function writeUint32(buf: Uint8Array, offset: number, value: number) {
  buf[offset] = value & 0xff;
  buf[offset + 1] = (value >>> 8) & 0xff;
  buf[offset + 2] = (value >>> 16) & 0xff;
  buf[offset + 3] = (value >>> 24) & 0xff;
}

function concatArrays(parts: Uint8Array[]) {
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

export function buildZip(files: ZipFile[]) {
  const encoder = new TextEncoder();
  const fileParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let offset = 0;

  for (const file of files) {
    const path = file.path.replace(/\\/g, '/');
    const nameBytes = encoder.encode(path);
    const data = file.data;
    const crc = crc32(data);

    const local = new Uint8Array(30 + nameBytes.length);
    writeUint32(local, 0, 0x04034b50);
    writeUint16(local, 4, 20);
    writeUint16(local, 6, 0);
    writeUint16(local, 8, 0);
    writeUint16(local, 10, 0);
    writeUint16(local, 12, 0);
    writeUint32(local, 14, crc);
    writeUint32(local, 18, data.length);
    writeUint32(local, 22, data.length);
    writeUint16(local, 26, nameBytes.length);
    writeUint16(local, 28, 0);
    local.set(nameBytes, 30);

    fileParts.push(local, data);

    const central = new Uint8Array(46 + nameBytes.length);
    writeUint32(central, 0, 0x02014b50);
    writeUint16(central, 4, 20);
    writeUint16(central, 6, 20);
    writeUint16(central, 8, 0);
    writeUint16(central, 10, 0);
    writeUint16(central, 12, 0);
    writeUint16(central, 14, 0);
    writeUint32(central, 16, crc);
    writeUint32(central, 20, data.length);
    writeUint32(central, 24, data.length);
    writeUint16(central, 28, nameBytes.length);
    writeUint16(central, 30, 0);
    writeUint16(central, 32, 0);
    writeUint16(central, 34, 0);
    writeUint16(central, 36, 0);
    writeUint32(central, 38, 0);
    writeUint32(central, 42, offset);
    central.set(nameBytes, 46);

    centralParts.push(central);

    offset += local.length + data.length;
  }

  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const centralOffset = offset;

  const eocd = new Uint8Array(22);
  writeUint32(eocd, 0, 0x06054b50);
  writeUint16(eocd, 4, 0);
  writeUint16(eocd, 6, 0);
  writeUint16(eocd, 8, files.length);
  writeUint16(eocd, 10, files.length);
  writeUint32(eocd, 12, centralSize);
  writeUint32(eocd, 16, centralOffset);
  writeUint16(eocd, 20, 0);

  return concatArrays([...fileParts, ...centralParts, eocd]);
}
