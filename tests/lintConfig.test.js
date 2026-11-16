const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { fixture } = require('./helpers');
const { readFormatterOptions, tomlTemplate } = require('../out/lintConfig');

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
