const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { fixture, shortPath } = require('./helpers');
const { BridgeFiles, BridgeFileError, LINK_NAME, MASTER_NAME, fileLockState, defaultBridgeDir } = require('../out/bridgeFiles');

const readOnly = file => (fs.statSync(file).mode & 0o200) === 0;
function setup(t, content = 'bridge v2', options = {}) {
    const root = fixture(t);
    const bundledPath = path.join(root, 'ext', 'resources', 'bridge', 'C5D_Bridge.catvba');
    fs.mkdirSync(path.dirname(bundledPath), { recursive: true });
    fs.writeFileSync(bundledPath, content);
    const bridgeDir = path.join(root, 'local', 'Cat5Dev', 'bridge');
    const sha8 = createHash('sha256').update(content).digest('hex').slice(0, 8);
    const files = new BridgeFiles({ bundledPath, bridgeDir, ...options });
    // Read-only files would block the fixture cleanup.
    t.after(() => { for (const dir of [bridgeDir]) if (fs.existsSync(dir)) for (const name of fs.readdirSync(dir)) try { fs.chmodSync(path.join(dir, name), 0o666); } catch { } });
    return { root, bundledPath, bridgeDir, sha8, files, master: path.join(bridgeDir, `C5D_Bridge-${sha8}.catvba`) };
}
test('file name patterns follow the contract', () => {
    assert.ok(MASTER_NAME.test('C5D_Bridge-9d6a56f7.catvba'));
    assert.ok(!MASTER_NAME.test('C5D_Bridge-9d6a56f7-12345-20261121151114123.catvba'));
    for (const name of ['C5D_Bridge-9d6a56f7-12345-20261121151114123.catvba', 'C5D_Bridge-9d6a56f7-0-20261121151114123-1.catvba',
        'c5d_bridge-9D6A56F7-0-20261121151114123-99.CATVBA']) assert.ok(LINK_NAME.test(name), name);
    for (const name of ['C5D_Bridge-9d6a56f7.catvba', 'C5D_Bridge-9d6a56f7-0-2026112115111412.catvba', 'C5D_Bridge-9d6a56f7-0-20261121151114123-100.catvba',
        'C5D_Bridge-build-20261121-w.catvba', 'C5D_Bridge-9d6a56f7-x-20261121151114123.catvba']) assert.ok(!LINK_NAME.test(name), name);
    assert.match(defaultBridgeDir(), /[\\/]Cat5Dev[\\/]bridge$/);
});
test('prepare copies the bundled bridge to the hash-named read-only master, then reuses it', t => {
    const { files, master, bundledPath } = setup(t);
    const first = files.prepare();
    assert.equal(first.master, master);
    assert.ok(path.isAbsolute(first.master));
    assert.equal(first.action, 'copied');
    assert.equal(fs.readFileSync(master, 'utf8'), 'bridge v2');
    assert.ok(readOnly(master));
    assert.ok(!readOnly(bundledPath), 'the bundled file is left as is');
    const second = files.prepare();
    assert.equal(second.action, 'reused'); assert.equal(second.master, master);
    // A master that lost its read-only attribute gets it back.
    fs.chmodSync(master, 0o666);
    assert.equal(files.prepare().action, 'reused');
    assert.ok(readOnly(master));
});
test('a damaged master is replaced when unlocked and stops the operation when locked', t => {
    const unlocked = setup(t);
    fs.mkdirSync(unlocked.bridgeDir, { recursive: true });
    fs.writeFileSync(unlocked.master, 'half a cop'); fs.chmodSync(unlocked.master, 0o444);
    const prepared = unlocked.files.prepare();
    assert.equal(prepared.action, 'replaced');
    assert.equal(fs.readFileSync(unlocked.master, 'utf8'), 'bridge v2'); assert.ok(readOnly(unlocked.master));
    const locked = setup(t, 'bridge v2', { lockState: () => 'locked' });
    fs.mkdirSync(locked.bridgeDir, { recursive: true });
    fs.writeFileSync(locked.master, 'damaged');
    assert.throws(() => locked.files.prepare(), e => e instanceof BridgeFileError && /restart CATIA/.test(e.message));
    assert.equal(fs.readFileSync(locked.master, 'utf8'), 'damaged');
});
test('a missing bundled bridge or a master that is not a file is a BridgeFileError', t => {
    const missing = setup(t);
    fs.unlinkSync(missing.bundledPath);
    assert.throws(() => missing.files.prepare(), BridgeFileError);
    const directory = setup(t);
    fs.mkdirSync(directory.master, { recursive: true });
    assert.throws(() => directory.files.prepare(), BridgeFileError);
});
test('cleanup deletes unlocked links, copies and old masters, keeps locked ones, and leaves other names alone', t => {
    const locked = new Set();
    const { files, bridgeDir, master, sha8 } = setup(t, 'bridge v2', { lockState: file => locked.has(path.basename(file)) ? 'locked' : 'unlocked' });
    files.prepare();
    const stamp = '20261121151114123';
    const hardLink = `C5D_Bridge-${sha8}-4321-${stamp}.catvba`;
    const copy = `C5D_Bridge-${sha8}-0-${stamp}-1.catvba`;
    const lockedLink = `C5D_Bridge-${sha8}-4321-${stamp}-2.catvba`;
    const oldMaster = 'C5D_Bridge-00000000.catvba';
    const oldLink = `C5D_Bridge-00000000-4321-${stamp}.catvba`;
    const lockedOldMaster = 'C5D_Bridge-11111111.catvba';
    const unrelated = ['C5D_Bridge-build-20261121-w.catvba', 'C5D_Bridge-probe.catvba', 'notes.txt', `C5D_Bridge-${sha8}-4321-${stamp}.catvba.tmp`];
    fs.linkSync(master, path.join(bridgeDir, hardLink));
    fs.linkSync(master, path.join(bridgeDir, lockedLink));
    for (const name of [copy, oldMaster, oldLink, lockedOldMaster, ...unrelated]) { fs.writeFileSync(path.join(bridgeDir, name), name); }
    for (const name of [copy, oldMaster, oldLink, lockedOldMaster]) fs.chmodSync(path.join(bridgeDir, name), 0o444);
    fs.mkdirSync(path.join(bridgeDir, `C5D_Bridge-${sha8}-1-${stamp}.catvba`)); // A directory is never touched.
    locked.add(lockedLink); locked.add(lockedOldMaster);
    const prepared = files.prepare();
    const left = new Set(fs.readdirSync(bridgeDir));
    for (const name of [hardLink, copy, oldMaster, oldLink]) assert.ok(!left.has(name), `${name} deleted`);
    for (const name of [lockedLink, lockedOldMaster, ...unrelated, path.basename(master), `C5D_Bridge-${sha8}-1-${stamp}.catvba`]) assert.ok(left.has(name), `${name} kept`);
    // Deleting a hard link cleared the shared read-only attribute; it is restored on the master and on kept files.
    assert.ok(readOnly(master)); assert.ok(readOnly(path.join(bridgeDir, lockedLink))); assert.ok(readOnly(path.join(bridgeDir, lockedOldMaster)));
    assert.ok(!readOnly(path.join(bridgeDir, 'notes.txt')));
    const actions = Object.fromEntries(prepared.cleanup.map(e => [path.basename(e.file), e.action]));
    assert.deepEqual(actions, { [hardLink]: 'deleted', [copy]: 'deleted', [oldMaster]: 'deleted', [oldLink]: 'deleted', [lockedLink]: 'kept (locked)', [lockedOldMaster]: 'kept (locked)' });
});
test('a file another process has open is reported as locked and kept', { skip: process.platform !== 'win32' }, t => {
    const { files, bridgeDir, master, sha8 } = setup(t);
    files.prepare();
    const link = path.join(bridgeDir, `C5D_Bridge-${sha8}-1-20261121151114123.catvba`);
    fs.linkSync(master, link);
    const fd = fs.openSync(link, 'r');
    try {
        assert.equal(fileLockState(link), 'locked');
        const prepared = files.prepare();
        assert.ok(fs.existsSync(link));
        assert.deepEqual(prepared.cleanup.map(e => e.action), ['kept (locked)']);
    } finally { fs.closeSync(fd); }
    assert.equal(fileLockState(link), 'unlocked');
    files.prepare();
    assert.ok(!fs.existsSync(link)); assert.ok(readOnly(master));
});
test('an unknown lock state keeps the file', t => {
    const { files, bridgeDir, sha8 } = setup(t, 'bridge v2', { lockState: () => 'unknown' });
    fs.mkdirSync(bridgeDir, { recursive: true });
    const link = path.join(bridgeDir, `C5D_Bridge-${sha8}-1-20261121151114123.catvba`);
    fs.writeFileSync(link, 'x');
    const prepared = files.prepare();
    assert.ok(fs.existsSync(link));
    assert.deepEqual(prepared.cleanup.map(e => e.action), ['kept (lock unknown)']);
});
test('the master path uses the long name of the bridge folder (fs.realpathSync.native expands 8.3 short names)', t => {
    const { files, bridgeDir, master, bundledPath } = setup(t);
    fs.mkdirSync(bridgeDir, { recursive: true });
    const calls = [];
    const native = fs.realpathSync.native;
    fs.realpathSync.native = (file, ...rest) => { calls.push(file); return native(file, ...rest); };
    const expected = path.join(native(bridgeDir), path.basename(master));
    try { assert.equal(files.prepare().master, expected); }
    finally { fs.realpathSync.native = native; }
    assert.ok(calls.includes(bridgeDir), JSON.stringify(calls));
    const short = shortPath(bridgeDir);
    if (!short) { t.diagnostic('no 8.3 short name on this volume; the short-path half is skipped'); return; }
    const viaShort = new BridgeFiles({ bundledPath, bridgeDir: short });
    const prepared = viaShort.prepare();
    assert.equal(prepared.master, expected, `${short} -> ${prepared.master}`);
    assert.equal(viaShort.masterPath(), expected);
});
