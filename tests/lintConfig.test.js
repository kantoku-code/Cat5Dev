const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { fixture } = require('./helpers');
const {
    readFormatterOptions, readLintOptions, readProjectSettings, tomlTemplate, writeTomlKey, writeTomlProjectKey,
} = require('../out/lintConfig');

function writeToml(t, content) {
    const root = fixture(t);
    if (content !== undefined) fs.writeFileSync(path.join(root, 'cat5dev.toml'), content);
    return root;
}

test('formatter is disabled without cat5dev.toml, without [formatter], or with enabled = false', t => {
    for (const content of [undefined, '[project]\nlanguage = "ja"\n', '[formatter]\nenabled = false\nformat_on_save = true\n']) {
        const opts = readFormatterOptions(writeToml(t, content));
        assert.equal(opts.enabled, false);
    }
});

test('enabling the formatter alone does not format on save or on pull', t => {
    const opts = readFormatterOptions(writeToml(t, '[formatter]\nenabled = true\n'));
    assert.equal(opts.enabled, true);
    assert.equal(opts.format_on_save, false);
    assert.equal(opts.format_on_pull, undefined);
    assert.equal(opts.indent_size, 4);
});

test('format_on_save and format_on_pull are read from [formatter]', t => {
    const opts = readFormatterOptions(writeToml(t, '[formatter]\nenabled = true\nformat_on_save = true\nformat_on_pull = TRUE\nindent_size = 2\n'));
    assert.equal(opts.format_on_save, true);
    assert.equal(opts.format_on_pull, true);
    assert.equal(opts.indent_size, 2);
    assert.equal(readFormatterOptions(writeToml(t, '[formatter]\nenabled = true\nformat_on_pull = false\n')).format_on_pull, false);
});

test('generated template writes format_on_save = false and format_on_pull = false for every language', t => {
    for (const lang of ['ja', 'en', 'zh']) {
        const text = tomlTemplate(lang);
        assert.match(text, /^format_on_save = false$/m);
        assert.match(text, /^format_on_pull = false$/m);
        const opts = readFormatterOptions(writeToml(t, text.replace(/^\[formatter\]\n(#[^\n]*\n)*enabled = false/m, m => m.replace('enabled = false', 'enabled = true'))));
        assert.equal(opts.enabled, true, lang);
        assert.equal(opts.format_on_save, false, lang);
        assert.equal(opts.format_on_pull, false, lang);
    }
});

test('generated template no longer writes the transfer encoding', () => {
    for (const lang of ['ja', 'en', 'zh']) {
        const text = tomlTemplate(lang);
        assert.doesNotMatch(text, /^encoding\s*=/m, lang);
        assert.doesNotMatch(text, /shift_jis|gbk/i, lang);
        assert.match(text, new RegExp(`^language = "${lang}"$`, 'm'));
    }
});

test('a leftover [project] encoding is ignored without error', t => {
    const root = writeToml(t, '[project]\ntarget_project = "Review"\nlanguage = "ja"\nencoding = "shift_jis"\n');
    assert.deepEqual(readProjectSettings(root), { targetProject: 'Review', language: 'ja' });
    assert.deepEqual(readProjectSettings(writeToml(t, undefined)), { targetProject: '', language: '' });
});

// ---------- writeTomlKey ----------

const readToml = root => fs.readFileSync(path.join(root, 'cat5dev.toml'), 'utf-8');

test('writeTomlKey replaces only the key in the given section and keeps comments and other lines', t => {
    const root = writeToml(t, [
        '# header',
        '[project]',
        'ansi_code_page = 1',
        '[lint.rules]',
        '# comment (0 = auto)',
        'ansi_code_page = 0',
        'goto = false',
        '',
    ].join('\n'));
    writeTomlKey(root, 'lint.rules', 'ansi_code_page', '1252');
    assert.equal(readToml(root), [
        '# header',
        '[project]',
        'ansi_code_page = 1',
        '[lint.rules]',
        '# comment (0 = auto)',
        'ansi_code_page = 1252',
        'goto = false',
        '',
    ].join('\n'));
});

test('writeTomlKey adds the key right after the section header when the key is missing', t => {
    const root = writeToml(t, '[lint]\nenabled = true\n[lint.rules]\ngoto = false\n[formatter]\nenabled = false\n');
    writeTomlKey(root, 'lint.rules', 'ansi_code_page', '936');
    assert.equal(readToml(root), '[lint]\nenabled = true\n[lint.rules]\nansi_code_page = 936\ngoto = false\n[formatter]\nenabled = false\n');
    assert.equal(readLintOptions(root).ansi_code_page, 936);
});

test('writeTomlKey appends the section when it is missing, and does nothing without cat5dev.toml', t => {
    const root = writeToml(t, '[lint]\nenabled = true\n\n');
    writeTomlKey(root, 'lint.rules', 'ansi_code_page', '1252');
    assert.equal(readToml(root), '[lint]\nenabled = true\n\n[lint.rules]\nansi_code_page = 1252\n');
    assert.equal(readLintOptions(root).ansi_code_page, 1252);

    const empty = writeToml(t, undefined);
    writeTomlKey(empty, 'lint.rules', 'ansi_code_page', '1252');
    assert.equal(fs.existsSync(path.join(empty, 'cat5dev.toml')), false);
});

test('writeTomlProjectKey still writes a quoted string into [project]', t => {
    const root = writeToml(t, '[project]\ntarget_project = ""\nlanguage = "ja"\n');
    writeTomlProjectKey(root, 'language', 'en');
    assert.deepEqual(readProjectSettings(root), { targetProject: '', language: 'en' });
    assert.match(readToml(root), /^language = "en"$/m);
});

test('switching the template language and the code page together gives the expected pair', t => {
    for (const [from, to, codePage] of [['ja', 'en', 1252], ['en', 'zh', 936], ['zh', 'ja', 932]]) {
        const root = writeToml(t, tomlTemplate(from).replace(/^\[lint\]\nenabled = false/m, '[lint]\nenabled = true'));
        writeTomlProjectKey(root, 'language', to);
        writeTomlKey(root, 'lint.rules', 'ansi_code_page', String(codePage));
        assert.equal(readProjectSettings(root).language, to);
        assert.equal(readLintOptions(root).ansi_code_page, codePage, `${from} -> ${to}`);
        assert.equal((readToml(root).match(/^ansi_code_page = /gm) ?? []).length, 1);
    }
});
