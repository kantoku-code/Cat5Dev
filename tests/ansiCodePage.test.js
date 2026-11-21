const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const { fixture } = require('./helpers');
const {
    AnsiCharChecker, getAnsiCharChecker, findNonAnsiRanges, iconvNameForCodePage, SUPPORTED_CODE_PAGES,
    estimateCodePageFromLanguage, parseAcpFromRegOutput, detectSystemAnsiCodePage, NON_ANSI_DIAGNOSTIC_CODE,
} = require('../out/ansiCodePage');
const { readLintOptions, tomlTemplate } = require('../out/lintConfig');

const cp932 = getAnsiCharChecker(932);
const cp936 = getAnsiCharChecker(936);

// ---------- 判定（findNonAnsiRanges） ----------

test('ASCII-only text has no diagnostics', () => {
    const text = "Option Explicit\r\nSub Main()\r\n    MsgBox \"Hello\" ' comment\r\nEnd Sub\r\n";
    assert.deepEqual(findNonAnsiRanges(text, cp932), []);
    assert.deepEqual(findNonAnsiRanges(text, getAnsiCharChecker(1252)), []);
});

test('an emoji is reported with a UTF-16 range that covers the whole surrogate pair', () => {
    const line = 'MsgBox "OK🙂"';
    const ranges = findNonAnsiRanges(line, cp932);
    assert.deepEqual(ranges, [{ line: 0, startCol: 10, endCol: 12, text: '🙂' }]);
    assert.equal(line.slice(10, 12), '🙂');
});

test('Japanese text including ① and half-width katakana round-trips in CP932', () => {
    assert.deepEqual(findNonAnsiRanges("Dim 日本語① As String ' 日本語①ｱ\r\nMsgBox \"ｱｲｳ髙㈱\"", cp932), []);
});

test('half-width katakana is reported in CP936 but kanji is not', () => {
    const ranges = findNonAnsiRanges("' 日本語①ｱ", cp936);
    assert.deepEqual(ranges, [{ line: 0, startCol: 6, endCol: 7, text: 'ｱ' }]);
});

test('consecutive bad characters are merged; separated ones are reported separately', () => {
    const text = "x = \"🙂🙂\" ' 日本🎉語 🙂\nMsgBox 1";
    const ranges = findNonAnsiRanges(text, cp932);
    assert.deepEqual(ranges.map(r => r.text), ['🙂🙂', '🎉', '🙂']);
    for (const r of ranges) {
        assert.equal(text.split('\n')[r.line].slice(r.startCol, r.endCol), r.text);
    }
});

test('characters in identifiers, strings and comments are all checked, on the right lines', () => {
    const text = ['Dim a🙂 As Long', 'MsgBox "🙂"', "' 🙂", 'End Sub'].join('\r\n');
    const ranges = findNonAnsiRanges(text, cp932);
    assert.deepEqual(ranges.map(r => [r.line, r.startCol, r.endCol]), [[0, 5, 7], [1, 8, 10], [2, 2, 4]]);
});

test('characters that Windows maps to a different character are reported (〜 U+301C, ¥, NBSP)', () => {
    const ranges = findNonAnsiRanges('a〜b¥c d', cp932);
    assert.deepEqual(ranges.map(r => r.text), ['〜', '¥', ' ']);
    // 全角チルダ ～ U+FF5E は CP932 で往復する
    assert.deepEqual(findNonAnsiRanges('～', cp932), []);
});

test('lone surrogates are reported', () => {
    const ranges = findNonAnsiRanges('a\uD83Db', cp932);
    assert.deepEqual(ranges, [{ line: 0, startCol: 1, endCol: 2, text: '\uD83D' }]);
});

test('single-byte code pages: é is fine in 1252 but 日 is not', () => {
    const c = getAnsiCharChecker(1252);
    assert.deepEqual(findNonAnsiRanges('café 日', c).map(r => r.text), ['日']);
});

test('all listed code pages are supported and unknown ones are not', () => {
    for (const cp of [874, 932, 936, 949, 950, 1250, 1251, 1252, 1253, 1254, 1255, 1256, 1257, 1258]) {
        assert.ok(SUPPORTED_CODE_PAGES.includes(cp), String(cp));
        assert.ok(iconvNameForCodePage(cp), String(cp));
        assert.ok(AnsiCharChecker.create(cp), String(cp));
    }
    for (const cp of [0, 437, 65001, 12345]) {
        assert.equal(getAnsiCharChecker(cp), undefined, String(cp));
    }
});

test('a large file is checked quickly', () => {
    const lines = [];
    for (let i = 0; i < 50000; i++) {
        lines.push(i % 10 === 0 ? `    x = "日本語の文字列 ${i}" ' コメント①` : `    y = y + ${i} ' ascii only`);
    }
    lines.push('MsgBox "🙂"');
    const start = process.hrtime.bigint();
    const ranges = findNonAnsiRanges(lines.join('\r\n'), cp932);
    const ms = Number(process.hrtime.bigint() - start) / 1e6;
    assert.equal(ranges.length, 1);
    assert.ok(ms < 2000, `took ${ms} ms`);
});

// ---------- コードページの検出 ----------

test('ACP is parsed from reg.exe output', () => {
    const out = '\r\nHKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\Nls\\CodePage\r\n    ACP    REG_SZ    932\r\n\r\n';
    assert.equal(parseAcpFromRegOutput(out), 932);
    assert.equal(parseAcpFromRegOutput('    OEMCP    REG_SZ    437\r\n'), undefined);
    assert.equal(parseAcpFromRegOutput(''), undefined);
});

test('detection uses the registry value, and falls back to the display language on failure', async () => {
    assert.deepEqual(await detectSystemAnsiCodePage('ja', async () => '    ACP    REG_SZ    936\r\n'),
        { codePage: 936, source: 'registry' });
    const failed = await detectSystemAnsiCodePage('zh-cn', async () => { throw new Error('reg failed'); });
    assert.equal(failed.codePage, 936);
    assert.equal(failed.source, 'fallback');
    assert.equal(failed.reason, 'reg failed');
    assert.equal((await detectSystemAnsiCodePage('en', async () => 'nothing')).codePage, 932);
});

test('fallback code page is estimated from the display language', () => {
    assert.equal(estimateCodePageFromLanguage('ja'), 932);
    assert.equal(estimateCodePageFromLanguage('zh-cn'), 936);
    assert.equal(estimateCodePageFromLanguage('zh-tw'), 950);
    assert.equal(estimateCodePageFromLanguage('ko'), 949);
    assert.equal(estimateCodePageFromLanguage('en'), 932);
    assert.equal(estimateCodePageFromLanguage(undefined), 932);
});

// ---------- 設定（cat5dev.toml） ----------

function writeToml(t, content) {
    const root = fixture(t);
    if (content !== undefined) fs.writeFileSync(path.join(root, 'cat5dev.toml'), content);
    return root;
}

test('non_ansi_character is on by default once lint is enabled, and off when lint is disabled', t => {
    const on = readLintOptions(writeToml(t, '[lint]\nenabled = true\n'));
    assert.equal(on.non_ansi_character, true);
    assert.equal(on.ansi_code_page, 0);
    assert.equal(readLintOptions(writeToml(t, '[lint]\nenabled = false\n')).non_ansi_character, false);
    assert.equal(readLintOptions(writeToml(t, undefined)).non_ansi_character, false);
});

test('non_ansi_character can be disabled and ansi_code_page can be set', t => {
    const opts = readLintOptions(writeToml(t, '[lint]\nenabled = true\n[lint.rules]\nnon_ansi_character = false\nansi_code_page = 936\n'));
    assert.equal(opts.non_ansi_character, false);
    assert.equal(opts.ansi_code_page, 936);
    for (const [value, expected] of [['"936"', 936], ['"cp949"', 949], ['"auto"', 0], ['0', 0], ['"bogus"', 0]]) {
        const o = readLintOptions(writeToml(t, `[lint]\nenabled = true\n[lint.rules]\nansi_code_page = ${value}\n`));
        assert.equal(o.ansi_code_page, expected, value);
    }
});

test('generated template contains non_ansi_character and ansi_code_page for every language', t => {
    for (const lang of ['ja', 'en', 'zh']) {
        const text = tomlTemplate(lang);
        assert.match(text, /^non_ansi_character = true$/m, lang);
        assert.match(text, /^ansi_code_page = 0$/m, lang);
        const opts = readLintOptions(writeToml(t, text.replace(/^\[lint\]\nenabled = false/m, '[lint]\nenabled = true')));
        assert.equal(opts.non_ansi_character, true, lang);
        assert.equal(opts.ansi_code_page, 0, lang);
        assert.equal(opts.unmatched_blocks, true, lang);
    }
});

// ---------- Linter への組み込み（vscode をモックする） ----------

function loadLinter(root, language = 'ja') {
    class Range { constructor(sl, sc, el, ec) { Object.assign(this, { start: { line: sl, character: sc }, end: { line: el, character: ec } }); } }
    class Diagnostic { constructor(range, message, severity) { Object.assign(this, { range, message, severity }); } }
    const store = new Map();
    const vscode = {
        Range, Diagnostic,
        DiagnosticSeverity: { Error: 0, Warning: 1 },
        env: { language },
        languages: {
            createDiagnosticCollection: () => ({
                set: (uri, d) => store.set(uri.toString(), d), get: uri => store.get(uri.toString()),
                delete: uri => store.delete(uri.toString()), dispose() { },
            }),
        },
        workspace: {
            workspaceFolders: [{ uri: { fsPath: root } }],
            getConfiguration: () => ({ inspect: () => undefined }),
        },
    };
    const originalLoad = Module._load;
    Module._load = function (id, ...args) { return id === 'vscode' ? vscode : originalLoad.call(this, id, ...args); };
    for (const m of ['linter', 'i18n', 'lintConfig', 'vbaServer']) {
        delete require.cache[require.resolve(`../out/${m}`)];
    }
    try {
        return { ...require('../out/linter'), store };
    } finally {
        Module._load = originalLoad;
    }
}

function fakeDocument(text) {
    const lines = text.split(/\r\n|\n/);
    return {
        uri: { toString: () => 'file:///m.bas_utf' },
        getText: () => text,
        lineCount: lines.length,
        lineAt: i => ({ text: lines[i] }),
    };
}

const noServer = { getBaseUrl: () => null };

test('the linter reports VBA010 with a localized message for the configured code page', async t => {
    const root = writeToml(t, '[project]\nlanguage = "en"\n[lint]\nenabled = true\n[lint.rules]\nansi_code_page = 936\n');
    const { VbaLinter, store } = loadLinter(root);
    const log = [];
    const linter = new VbaLinter(noServer, { appendLine: l => log.push(l) });
    await linter.lintNow(fakeDocument("Sub Main()\r\n    MsgBox \"日本ｱ🙂\"\r\nEnd Sub"));
    const diags = store.get('file:///m.bas_utf');
    assert.equal(diags.length, 1);
    assert.equal(diags[0].code, NON_ANSI_DIAGNOSTIC_CODE);
    assert.equal(diags[0].source, 'cat5dev');
    assert.equal(diags[0].severity, 1);
    assert.deepEqual([diags[0].range.start, diags[0].range.end], [{ line: 1, character: 14 }, { line: 1, character: 17 }]);
    assert.equal(diags[0].message, 'CATIA VBA (code page 936) cannot represent this character, so Push replaces it with ? or another character: ｱ🙂');
    linter.dispose();
});

test('the linter logs and skips VBA010 for an unsupported code page, and honours non_ansi_character = false', async t => {
    const root = writeToml(t, '[project]\nlanguage = "ja"\n[lint]\nenabled = true\n[lint.rules]\nansi_code_page = 65001\n');
    const { VbaLinter, store } = loadLinter(root);
    const log = [];
    const linter = new VbaLinter(noServer, { appendLine: l => log.push(l) });
    await linter.lintNow(fakeDocument('MsgBox "🙂"'));
    await linter.lintNow(fakeDocument('MsgBox "🙂"'));
    assert.deepEqual(store.get('file:///m.bas_utf'), []);
    assert.equal(log.filter(l => l.includes('65001')).length, 1);

    fs.writeFileSync(path.join(root, 'cat5dev.toml'), '[lint]\nenabled = true\n[lint.rules]\nnon_ansi_character = false\nansi_code_page = 932\n');
    await linter.lintNow(fakeDocument('MsgBox "🙂"'));
    assert.deepEqual(store.get('file:///m.bas_utf'), []);
    linter.dispose();
});

test('the linter logs the detected system code page once', async t => {
    const root = writeToml(t, '[lint]\nenabled = true\n');
    const { VbaLinter } = loadLinter(root);
    const log = [];
    const linter = new VbaLinter(noServer, { appendLine: l => log.push(l) });
    await linter.lintNow(fakeDocument('MsgBox "🙂"'));
    await linter.lintNow(fakeDocument('MsgBox "🙂"'));
    assert.equal(log.filter(l => l.includes('ANSI')).length, 1, log.join('\n'));
    linter.dispose();
});
