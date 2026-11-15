const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { fixture } = require('./helpers');
const { captureModules, applyPull, listPullDeletions, PullConflictError, PullRecoveryError } = require('../out/pullTransaction');
function setup(t) {
    const root = fixture(t), dir = path.join(root, 'modules'); fs.mkdirSync(dir);
    fs.writeFileSync(path.join(dir, 'A.bas_utf'), 'old A'); fs.writeFileSync(path.join(dir, 'B.cls_utf'), 'old B');
    fs.writeFileSync(path.join(dir, 'notes.txt'), 'keep'); fs.mkdirSync(path.join(dir, 'subdir'));
    return { root, dir, snapshot: captureModules(root), incoming: new Map([['A.bas_utf', 'new A'], ['C.frm_utf', '']]) };
}
test('Pull keep/delete choices preserve non-managed paths and empty modules', t => {
    for (const remove of [false, true]) {
        const { root, dir, snapshot, incoming } = setup(t);
        assert.deepEqual(listPullDeletions(snapshot, incoming), ['B.cls_utf']);
        assert.deepEqual(applyPull(root, snapshot, incoming, remove), { writtenCount: 2, deletedCount: remove ? 1 : 0 });
        assert.equal(fs.readFileSync(path.join(dir, 'A.bas_utf'), 'utf8'), 'new A');
        assert.equal(fs.readFileSync(path.join(dir, 'C.frm_utf'), 'utf8'), '');
        assert.equal(fs.existsSync(path.join(dir, 'B.cls_utf')), !remove);
        assert.equal(fs.readFileSync(path.join(dir, 'notes.txt'), 'utf8'), 'keep');
        assert.ok(fs.statSync(path.join(dir, 'subdir')).isDirectory());
        assert.equal(fs.readdirSync(root).filter(n => n.startsWith('.cat5dev-pull-')).length, 0);
    }
});
test('valid zero-module Pull deletes only when confirmed', t => {
    for (const remove of [false, true]) {
        const { root, dir, snapshot } = setup(t);
        applyPull(root, snapshot, new Map(), remove);
        assert.equal(fs.existsSync(path.join(dir, 'A.bas_utf')), !remove);
        assert.ok(fs.existsSync(path.join(dir, 'notes.txt')));
    }
});
test('local modification, addition, deletion and dirty documents abort without applying', t => {
    for (const change of ['modify', 'add', 'delete', 'dirty']) {
        const { root, dir, snapshot, incoming } = setup(t);
        if (change === 'modify') fs.writeFileSync(path.join(dir, 'A.bas_utf'), 'edited');
        if (change === 'add') fs.writeFileSync(path.join(dir, 'D.bas_utf'), 'new');
        if (change === 'delete') fs.unlinkSync(path.join(dir, 'B.cls_utf'));
        const before = fs.readdirSync(dir).filter(n => fs.statSync(path.join(dir, n)).isFile()).map(n => [n, fs.readFileSync(path.join(dir, n), 'utf8')]);
        assert.throws(() => applyPull(root, snapshot, incoming, true, { isDirty: () => change === 'dirty' }), PullConflictError);
        assert.deepEqual(fs.readdirSync(dir).filter(n => fs.statSync(path.join(dir, n)).isFile()).map(n => [n, fs.readFileSync(path.join(dir, n), 'utf8')]), before);
    }
});
test('commit failure restores overwritten, newly created and deleted files', t => {
    const { root, dir, snapshot, incoming } = setup(t);
    assert.throws(() => applyPull(root, snapshot, incoming, true, { beforeMutation(phase, name) { if (phase === 'apply' && name === 'B.cls_utf') throw Error('disk failure'); } }), /disk failure/);
    assert.equal(fs.readFileSync(path.join(dir, 'A.bas_utf'), 'utf8'), 'old A');
    assert.equal(fs.readFileSync(path.join(dir, 'B.cls_utf'), 'utf8'), 'old B');
    assert.ok(!fs.existsSync(path.join(dir, 'C.frm_utf')));
});
test('restore failure retains backup and recovery manifest', t => {
    const { root, snapshot, incoming } = setup(t);
    assert.throws(() => applyPull(root, snapshot, incoming, true, { beforeMutation(phase, name) {
        if (phase === 'apply' && name === 'B.cls_utf' || phase === 'rollback' && name === 'A.bas_utf') throw Error('injected');
    } }), e => {
        assert.ok(e instanceof PullRecoveryError);
        assert.ok(e.unrestored.includes('A.bas_utf'));
        assert.equal(fs.readFileSync(path.join(e.recoveryPath, 'backup', 'A.bas_utf'), 'utf8'), 'old A');
        assert.ok(fs.existsSync(path.join(e.recoveryPath, 'recovery.json'))); return true;
    });
});
test('new modules directory is removed on failed transaction', t => {
    const root = fixture(t), snapshot = captureModules(root);
    assert.throws(() => applyPull(root, snapshot, new Map([['A.bas_utf', 'A'], ['B.bas_utf', 'B']]), true,
        { beforeMutation(phase, name) { if (phase === 'apply' && name === 'B.bas_utf') throw Error('injected'); } }));
    assert.ok(!fs.existsSync(path.join(root, 'modules')));
});
test('incoming traversal and case collisions are rejected', t => {
    const { root, snapshot } = setup(t);
    for (const incoming of [new Map([['../bad.bas_utf', 'x']]), new Map([['A.bas_utf', '1'], ['a.bas_utf', '2']])]) {
        assert.throws(() => applyPull(root, snapshot, incoming, true), PullConflictError);
    }
});
