// Reads VBA module sources out of a .catvba (an OLE compound file holding a VBA project) without CATIA (developer tool, not shipped).
// Used by scripts/build-bridge-catvba.js to check offline that a rebuilt C5D_Bridge.catvba holds exactly the expected modules
// (MS-CFB: header, FAT, mini FAT, directory; MS-OVBA: the "dir" stream and the compressed source of each module stream).
// Only reads; any structural problem throws an Error whose message starts with "catvba:".
// Usage: node scripts/lib/catvba-read.js <file.catvba> [<module>]   (summary JSON, or the raw source of one module)
'use strict';
const fs = require('node:fs');

const SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const END_OF_CHAIN = 0xfffffffe, FREE = 0xffffffff, NO_STREAM = 0xffffffff;
const fail = message => { throw new Error(`catvba: ${message}`); };

/** Parses the compound file: { streams: Map<'VBA/dir', Buffer>, names } (storage path joined with '/', the root omitted). */
function readCompoundFile(buf) {
    if (buf.length < 512 || !buf.subarray(0, 8).equals(SIGNATURE)) fail('not an OLE compound file');
    const sectorSize = 1 << buf.readUInt16LE(30), miniSize = 1 << buf.readUInt16LE(32);
    if (![512, 4096].includes(sectorSize) || miniSize !== 64) fail('unexpected sector size');
    const fatSectors = buf.readUInt32LE(44), firstDir = buf.readUInt32LE(48), miniCutoff = buf.readUInt32LE(56);
    const firstMiniFat = buf.readUInt32LE(60), miniFatCount = buf.readUInt32LE(64);
    let difat = buf.readUInt32LE(68), difatCount = buf.readUInt32LE(72);
    const sector = n => {
        const start = (n + 1) * sectorSize;
        if (start + sectorSize > buf.length) fail(`sector ${n} outside the file`);
        return buf.subarray(start, start + sectorSize);
    };
    const fatIndex = [];
    for (let i = 0; i < 109 && fatIndex.length < fatSectors; i++) fatIndex.push(buf.readUInt32LE(76 + i * 4));
    while (fatIndex.length < fatSectors && difatCount-- > 0) {
        const s = sector(difat);
        for (let i = 0; i < sectorSize / 4 - 1 && fatIndex.length < fatSectors; i++) fatIndex.push(s.readUInt32LE(i * 4));
        difat = s.readUInt32LE(sectorSize - 4);
    }
    const fat = [];
    for (const n of fatIndex) { const s = sector(n); for (let i = 0; i < sectorSize / 4; i++) fat.push(s.readUInt32LE(i * 4)); }
    const chain = (start, table) => {
        const out = [];
        for (let n = start; n !== END_OF_CHAIN; n = table[n]) {
            if (n === FREE || n >= table.length || out.length > table.length) fail('broken sector chain');
            out.push(n);
        }
        return out;
    };
    const readChain = start => Buffer.concat(chain(start, fat).map(sector));
    const dir = readChain(firstDir);
    const entries = [];
    for (let off = 0; off + 128 <= dir.length; off += 128) {
        const nameLength = dir.readUInt16LE(off + 64);
        entries.push({
            name: nameLength >= 2 ? dir.toString('utf16le', off, off + nameLength - 2) : '',
            type: dir[off + 66], left: dir.readUInt32LE(off + 68), right: dir.readUInt32LE(off + 72), child: dir.readUInt32LE(off + 76),
            start: dir.readUInt32LE(off + 116), size: dir.readUInt32LE(off + 120)
        });
    }
    const root = entries[0];
    if (!root || root.type !== 5) fail('no root entry');
    const ministream = root.start === END_OF_CHAIN ? Buffer.alloc(0) : readChain(root.start).subarray(0, root.size);
    const miniFat = [];
    if (miniFatCount > 0) { const m = readChain(firstMiniFat); for (let i = 0; i + 4 <= m.length; i += 4) miniFat.push(m.readUInt32LE(i)); }
    const streams = new Map();
    const walk = (index, prefix, seen) => {
        if (index === NO_STREAM) return;
        if (seen.has(index) || index >= entries.length) fail('broken directory tree');
        seen.add(index);
        const e = entries[index];
        walk(e.left, prefix, seen); walk(e.right, prefix, seen);
        const full = prefix ? `${prefix}/${e.name}` : e.name;
        if (e.type === 2) {
            if (e.size < miniCutoff) {
                const parts = chain(e.start, miniFat).map(n => ministream.subarray(n * miniSize, (n + 1) * miniSize));
                streams.set(full, Buffer.concat(parts).subarray(0, e.size));
            } else streams.set(full, readChain(e.start).subarray(0, e.size));
        } else if (e.type === 1) walk(e.child, full, seen);
    };
    walk(root.child, '', new Set());
    return { streams, names: [...streams.keys()] };
}

/** MS-OVBA 2.4.1 decompression of a CompressedContainer. */
function decompress(buf) {
    if (!buf || buf[0] !== 1) fail('compressed container signature missing');
    const out = [];
    let pos = 1;
    while (pos + 2 <= buf.length) {
        const header = buf.readUInt16LE(pos);
        const size = (header & 0x0fff) + 3, compressed = header >> 15;
        if (((header >> 12) & 7) !== 3) fail('bad chunk signature');
        const end = Math.min(pos + size, buf.length);
        pos += 2;
        const chunkStart = out.length;
        if (!compressed) { for (let i = 0; i < 4096 && pos < end; i++) out.push(buf[pos++]); continue; }
        while (pos < end) {
            const flags = buf[pos++];
            for (let bit = 0; bit < 8 && pos < end; bit++) {
                if (!((flags >> bit) & 1)) { out.push(buf[pos++]); continue; }
                const token = buf.readUInt16LE(pos); pos += 2;
                const decompressed = out.length - chunkStart;
                let bitCount = 4;
                while ((1 << bitCount) < decompressed) bitCount++;
                const length = (token & (0xffff >> bitCount)) + 3, offset = (token >> (16 - bitCount)) + 1;
                if (offset > decompressed) fail('copy token before the chunk start');
                for (let k = 0; k < length; k++) out.push(out[out.length - offset]);
            }
        }
    }
    return Buffer.from(out);
}

/**
 * Parses the decompressed "dir" stream (MS-OVBA 2.3.4.2): { projectName, codePage, modules: [{ name, stream, offset, type }] }
 * (type: 'procedural' = MODULETYPE 0x0021 standard module, 'document' = 0x0022 class / form / document module).
 */
function parseDir(dir) {
    let pos = 0, codePage = 1252, projectName = null, current = null;
    const modules = [];
    while (pos + 6 <= dir.length) {
        const id = dir.readUInt16LE(pos);
        const size = id === 0x0009 ? 6 : dir.readUInt32LE(pos + 2); // PROJECTVERSION: Reserved (4) + 6 bytes of data
        const data = dir.subarray(pos + 6, pos + 6 + size);
        pos += 6 + size;
        if (id === 0x0003) codePage = data.readUInt16LE(0);
        else if (id === 0x0004) projectName = data.toString('latin1');
        else if (id === 0x0019) { current = { name: data.toString('latin1'), stream: null, offset: 0, type: null }; modules.push(current); }
        else if (id === 0x001a && current) current.stream = data.toString('latin1');
        else if (id === 0x0031 && current) current.offset = data.readUInt32LE(0);
        else if (id === 0x0021 && current) current.type = 'procedural';
        else if (id === 0x0022 && current) current.type = 'document';
        else if (id === 0x0010) break; // Terminator
    }
    return { projectName, codePage, modules };
}

/**
 * Reads a .catvba: { projectName, codePage, modules: [{ name, type, source }], streams } where source is the module text exactly as
 * stored (Attribute lines included, CRLF; bytes decoded as latin1, so ASCII sources round-trip unchanged).
 */
function readCatvba(file) {
    const { streams, names } = readCompoundFile(fs.readFileSync(file));
    const dirKey = names.find(n => /(^|\/)VBA\/dir$/i.test(n));
    if (!dirKey) fail('no VBA/dir stream');
    const base = dirKey.slice(0, -'dir'.length);
    const info = parseDir(decompress(streams.get(dirKey)));
    const modules = info.modules.map(m => {
        const stream = streams.get(base + (m.stream ?? m.name));
        if (!stream) fail(`module stream missing: ${m.name}`);
        return { name: m.name, type: m.type, source: decompress(stream.subarray(m.offset)).toString('latin1') };
    });
    return { projectName: info.projectName, codePage: info.codePage, modules, streams: names };
}

/** Module source without the leading Attribute lines (what CodeModule.Lines returns) and with a single trailing CRLF removed. */
function codeOf(source) {
    return String(source).replace(/^(?:Attribute [^\r\n]*\r\n)+/, '').replace(/(?:\r\n)+$/, '');
}

module.exports = { readCatvba, readCompoundFile, decompress, parseDir, codeOf };

if (require.main === module) {
    const file = process.argv[2];
    if (!file) { console.error('usage: node scripts/lib/catvba-read.js <file.catvba> [<module>]'); process.exit(2); }
    const r = readCatvba(file);
    const only = process.argv[3];
    if (only) {
        const m = r.modules.find(x => x.name === only);
        if (!m) { console.error(`no module ${only}`); process.exit(1); }
        process.stdout.write(m.source);
    } else {
        console.log(JSON.stringify({ projectName: r.projectName, codePage: r.codePage,
            modules: r.modules.map(m => ({ name: m.name, type: m.type, chars: m.source.length })), streams: r.streams }, null, 2));
    }
}
