import {inflateRawSync} from 'node:zlib';

const EOCD = 0x06054b50;
const ZIP64_EOCD = 0x06064b50;
const ZIP64_LOCATOR = 0x07064b50;
const CENTRAL = 0x02014b50;
const LOCAL = 0x04034b50;
const MAX_MEMBER_BYTES = 256 * 1024 * 1024;

export class PackageError extends Error {
  constructor(message, {part, cause} = {}) {
    super(part ? `${part}: ${message}` : message, {cause});
    this.name = 'PackageError';
    this.part = part;
  }
}

const crcTable = new Uint32Array(256);
for (let index = 0; index < 256; index++) {
  let value = index;
  for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  crcTable[index] = value >>> 0;
}

export function crc32(bytes) {
  let value = 0xffffffff;
  for (const byte of bytes) value = crcTable[(value ^ byte) & 0xff] ^ (value >>> 8);
  return (value ^ 0xffffffff) >>> 0;
}

function readUInt64LE(bytes, offset) {
  const value = bytes.readBigUInt64LE(offset);
  if (value > BigInt(Number.MAX_SAFE_INTEGER)) throw new PackageError('ZIP64 偏移超出安全范围');
  return Number(value);
}

function findEndOfCentralDirectory(bytes) {
  if (bytes.length < 22) return -1;
  const first = Math.max(0, bytes.length - 22 - 0xffff);
  for (let offset = bytes.length - 22; offset >= first; offset--) {
    if (bytes.readUInt32LE(offset) !== EOCD) continue;
    const commentLength = bytes.readUInt16LE(offset + 20);
    if (offset + 22 + commentLength === bytes.length) return offset;
  }
  return -1;
}

function readDirectory(bytes) {
  const end = findEndOfCentralDirectory(bytes);
  if (end < 0) throw new PackageError('缺少 ZIP 中央目录');
  if (bytes.readUInt16LE(end + 4) !== 0 || bytes.readUInt16LE(end + 6) !== 0) {
    throw new PackageError('不支持分卷 ZIP');
  }
  let count = bytes.readUInt16LE(end + 10);
  let size = bytes.readUInt32LE(end + 12);
  let offset = bytes.readUInt32LE(end + 16);
  if (count === 0xffff || size === 0xffffffff || offset === 0xffffffff) {
    const locator = end - 20;
    if (locator >= 0 && bytes.readUInt32LE(locator) === ZIP64_LOCATOR) {
      const zip64Offset = readUInt64LE(bytes, locator + 8);
      if (zip64Offset + 56 > bytes.length || bytes.readUInt32LE(zip64Offset) !== ZIP64_EOCD) {
        throw new PackageError('ZIP64 中央目录损坏');
      }
      count = readUInt64LE(bytes, zip64Offset + 32);
      size = readUInt64LE(bytes, zip64Offset + 40);
      offset = readUInt64LE(bytes, zip64Offset + 48);
    }
  }
  if (count < 0 || offset < 0 || size < 0 || offset + size > bytes.length) throw new PackageError('ZIP 中央目录越界');
  return {count, offset};
}

/**
 * Reads ZIP packaging lazily. Metadata is validated up front, but member bytes
 * are only inflated and CRC-checked when requested. A damaged member therefore
 * yields an error for that member without invalidating the rest of the package.
 */
export class ZipArchive {
  constructor(input) {
    const bytes = Buffer.isBuffer(input) ? input : Buffer.from(input);
    this.bytes = bytes;
    const {count, offset} = readDirectory(bytes);
    this._entries = new Map();
    let cursor = offset;
    for (let index = 0; index < count; index++) {
      if (cursor + 46 > bytes.length || bytes.readUInt32LE(cursor) !== CENTRAL) {
        throw new PackageError('ZIP 中央目录项损坏');
      }
      const flags = bytes.readUInt16LE(cursor + 8);
      const method = bytes.readUInt16LE(cursor + 10);
      const crc = bytes.readUInt32LE(cursor + 16);
      const compressedSize = bytes.readUInt32LE(cursor + 20);
      const uncompressedSize = bytes.readUInt32LE(cursor + 24);
      const nameSize = bytes.readUInt16LE(cursor + 28);
      const extraSize = bytes.readUInt16LE(cursor + 30);
      const noteSize = bytes.readUInt16LE(cursor + 32);
      const localOffset = bytes.readUInt32LE(cursor + 42);
      const end = cursor + 46 + nameSize + extraSize + noteSize;
      if (end > bytes.length) throw new PackageError('ZIP 中央目录项越界');
      const name = bytes.subarray(cursor + 46, cursor + 46 + nameSize).toString('utf8');
      if (!name || name.includes('\\') || name.startsWith('/') || name.split('/').includes('..')) {
        throw new PackageError(`不安全的 ZIP 成员名: ${name}`);
      }
      if (this._entries.has(name)) throw new PackageError(`ZIP 成员重复: ${name}`);
      // Per-member problems (encryption, compression, local header damage) are
      // deliberately deferred to read() so one damaged member cannot hide the
      // rest of the package.
      this._entries.set(name, {name, method, crc, compressedSize, uncompressedSize, localOffset, flags});
      cursor = end;
    }
  }

  names() { return [...this._entries.keys()]; }
  has(name) { return this._entries.has(name); }
  entry(name) { return this._entries.get(name); }

  read(name) {
    const entry = this._entries.get(name);
    if (!entry) throw new PackageError('包内不存在该资料', {part: name});
    if (entry.uncompressedSize > MAX_MEMBER_BYTES) throw new PackageError('成员过大', {part: name});
    if (entry.flags & 0x1) throw new PackageError('不支持加密的 ZIP 成员', {part: name});
    if (entry.method !== 0 && entry.method !== 8) throw new PackageError(`不支持的压缩类型 ${entry.method}`, {part: name});
    if (entry.localOffset + 30 > this.bytes.length || this.bytes.readUInt32LE(entry.localOffset) !== LOCAL) {
      throw new PackageError('本地文件头损坏', {part: name});
    }
    const localNameSize = this.bytes.readUInt16LE(entry.localOffset + 26);
    const localExtraSize = this.bytes.readUInt16LE(entry.localOffset + 28);
    const dataStart = entry.localOffset + 30 + localNameSize + localExtraSize;
    if (dataStart + entry.compressedSize > this.bytes.length) throw new PackageError('成员数据越界', {part: name});
    const compressed = this.bytes.subarray(dataStart, dataStart + entry.compressedSize);
    let content;
    try {
      content = entry.method === 0 ? Buffer.from(compressed) : inflateRawSync(compressed, {maxOutputLength: Math.max(1, entry.uncompressedSize)});
    } catch (cause) {
      throw new PackageError('解压失败', {part: name, cause});
    }
    if (content.length !== entry.uncompressedSize) throw new PackageError('成员长度与目录不一致', {part: name});
    if (crc32(content) !== entry.crc) throw new PackageError('成员校验和不一致', {part: name});
    return content;
  }
}
