import * as fs from 'fs';
import * as path from 'path';
import { Language } from './i18n';
import { codePageForLanguage } from './ansiCodePage';

export interface LintOptions {
    option_explicit: boolean;
    on_error_resume_next: boolean;
    goto: boolean;
    max_line_length: number;
    unused_variables: boolean;
    max_nesting_depth: number;
    max_function_lines: number;
    unmatched_parens: boolean;
    unmatched_blocks: boolean;
    /** VBA010: CATIA の VBA（ANSI コードページ）で表せない文字を警告する。Go の /lint には送らない */
    non_ansi_character: boolean;
    /** VBA010 の判定に使うコードページ。0 = システムの ANSI コードページを自動検出。Go の /lint には送らない */
    ansi_code_page: number;
}

export const DEFAULT_LINT_OPTIONS: LintOptions = {
    option_explicit: true,
    on_error_resume_next: false,
    goto: false,
    max_line_length: 100,
    unused_variables: true,
    max_nesting_depth: 5,
    max_function_lines: 300,
    unmatched_parens: true,
    unmatched_blocks: true,
    non_ansi_character: true,
    ansi_code_page: 0,
};

/** Lint 全体が無効（cat5dev.toml が無い、または [lint] enabled = false）のときの値 */
const DISABLED_LINT_OPTIONS: LintOptions = {
    option_explicit: false,
    on_error_resume_next: false,
    goto: false,
    max_line_length: 0,
    unused_variables: false,
    max_nesting_depth: 0,
    max_function_lines: 0,
    unmatched_parens: false,
    unmatched_blocks: false,
    non_ansi_character: false,
    ansi_code_page: 0,
};

/**
 * 最小 TOML パーサー
 * [section] と key = value（bool / int / string）に対応。
 * コメント（# 以降）を除去して解析する。
 */
function parseToml(content: string): Record<string, Record<string, string>> {
    const result: Record<string, Record<string, string>> = {};
    let section = '';

    for (let line of content.split(/\r?\n/)) {
        // コメント除去
        const commentIdx = line.indexOf('#');
        if (commentIdx >= 0) {
            line = line.substring(0, commentIdx);
        }
        line = line.trim();
        if (!line) { continue; }

        // セクション
        const secMatch = line.match(/^\[([^\]]+)\]$/);
        if (secMatch) {
            section = secMatch[1].trim();
            if (!result[section]) { result[section] = {}; }
            continue;
        }

        // key = value
        const eqIdx = line.indexOf('=');
        if (eqIdx < 0) { continue; }
        const key = line.substring(0, eqIdx).trim();
        const val = line.substring(eqIdx + 1).trim();
        if (!section) { continue; }
        if (!result[section]) { result[section] = {}; }
        result[section][key] = val;
    }

    return result;
}

function getBool(map: Record<string, string>, key: string, def: boolean): boolean {
    const v = map[key];
    if (v === undefined) { return def; }
    return v.toLowerCase() === 'true';
}

function getInt(map: Record<string, string>, key: string, def: number): number {
    const v = map[key];
    if (v === undefined) { return def; }
    const n = parseInt(v, 10);
    return isNaN(n) ? def : n;
}

/** コードページの値を読む。`932` / `"932"` / `"cp932"` / `"auto"`（= 0）を受け付ける。不正値は def */
function getCodePage(map: Record<string, string>, key: string, def: number): number {
    const v = map[key];
    if (v === undefined) { return def; }
    const s = v.replace(/^"|"$/g, '').trim().toLowerCase();
    if (s === 'auto') { return 0; }
    const m = /^(?:cp|windows-?)?(\d+)$/.exec(s);
    return m ? parseInt(m[1], 10) : def;
}

/** ワークスペースルートの cat5dev.toml を読み込んで LintOptions を返す */
export function readLintOptions(workspaceRoot: string): LintOptions {
    const tomlPath = path.join(workspaceRoot, 'cat5dev.toml');
    if (!fs.existsSync(tomlPath)) {
        return { ...DISABLED_LINT_OPTIONS };
    }

    let content: string;
    try {
        content = fs.readFileSync(tomlPath, 'utf-8');
    } catch {
        return { ...DEFAULT_LINT_OPTIONS };
    }

    const toml = parseToml(content);

    // [lint] セクションで enabled = false（デフォルト）なら全ルール無効
    const lintSection = toml['lint'] ?? {};
    if (getBool(lintSection, 'enabled', false) === false) {
        return { ...DISABLED_LINT_OPTIONS };
    }

    const rules = toml['lint.rules'] ?? {};
    return {
        option_explicit:     getBool(rules, 'option_explicit',     DEFAULT_LINT_OPTIONS.option_explicit),
        on_error_resume_next:getBool(rules, 'on_error_resume_next', DEFAULT_LINT_OPTIONS.on_error_resume_next),
        goto:                getBool(rules, 'goto',                 DEFAULT_LINT_OPTIONS.goto),
        max_line_length:     getInt (rules, 'max_line_length',      DEFAULT_LINT_OPTIONS.max_line_length),
        unused_variables:    getBool(rules, 'unused_variables',     DEFAULT_LINT_OPTIONS.unused_variables),
        max_nesting_depth:   getInt (rules, 'max_nesting_depth',    DEFAULT_LINT_OPTIONS.max_nesting_depth),
        max_function_lines:  getInt (rules, 'max_function_lines',   DEFAULT_LINT_OPTIONS.max_function_lines),
        unmatched_parens:    getBool(rules, 'unmatched_parens',     DEFAULT_LINT_OPTIONS.unmatched_parens),
        unmatched_blocks:    getBool(rules, 'unmatched_blocks',     DEFAULT_LINT_OPTIONS.unmatched_blocks),
        non_ansi_character:  getBool(rules, 'non_ansi_character',   DEFAULT_LINT_OPTIONS.non_ansi_character),
        ansi_code_page:      getCodePage(rules, 'ansi_code_page',   DEFAULT_LINT_OPTIONS.ansi_code_page),
    };
}

export interface FormatterOptions {
    enabled: boolean;
    indent_size: number;
    capitalize_keywords: boolean;
    fix_indentation: boolean;
    trim_trailing_space: boolean;
    ensure_continuation_space: boolean;
    indent_continuation_lines: boolean;
    max_blank_lines: number;
    normalize_operator_spacing: boolean;
    normalize_comma_spacing: boolean;
    normalize_comment_space: boolean;
    expand_type_suffixes: boolean;
    format_on_save: boolean;
    /** Pull で取り込んだファイルを整形する。toml にキーが無い場合は undefined */
    format_on_pull: boolean | undefined;
}

export const DEFAULT_FORMATTER_OPTIONS: FormatterOptions = {
    enabled: false,
    indent_size: 4,
    capitalize_keywords: true,
    fix_indentation: true,
    trim_trailing_space: true,
    ensure_continuation_space: true,
    indent_continuation_lines: true,
    max_blank_lines: 2,
    normalize_operator_spacing: false,
    normalize_comma_spacing: false,
    normalize_comment_space: false,
    expand_type_suffixes: false,
    format_on_save: false,
    format_on_pull: false,
};

/** ワークスペースルートの cat5dev.toml の [formatter] を読み込んで FormatterOptions を返す。
 *  toml・[formatter] セクションが無い場合、または enabled = false の場合は無効。 */
export function readFormatterOptions(workspaceRoot: string): FormatterOptions {
    const disabled: FormatterOptions = { ...DEFAULT_FORMATTER_OPTIONS, enabled: false, format_on_pull: undefined };
    const tomlPath = path.join(workspaceRoot, 'cat5dev.toml');
    if (!fs.existsSync(tomlPath)) { return disabled; }

    let content: string;
    try { content = fs.readFileSync(tomlPath, 'utf-8'); } catch { return disabled; }

    const sec = parseToml(content)['formatter'];
    if (!sec || getBool(sec, 'enabled', false) === false) { return disabled; }

    return {
        enabled: true,
        indent_size:               getInt (sec, 'indent_size',               DEFAULT_FORMATTER_OPTIONS.indent_size),
        capitalize_keywords:       getBool(sec, 'capitalize_keywords',       DEFAULT_FORMATTER_OPTIONS.capitalize_keywords),
        fix_indentation:           getBool(sec, 'fix_indentation',           DEFAULT_FORMATTER_OPTIONS.fix_indentation),
        trim_trailing_space:       getBool(sec, 'trim_trailing_space',       DEFAULT_FORMATTER_OPTIONS.trim_trailing_space),
        ensure_continuation_space: getBool(sec, 'ensure_continuation_space', DEFAULT_FORMATTER_OPTIONS.ensure_continuation_space),
        indent_continuation_lines: getBool(sec, 'indent_continuation_lines', DEFAULT_FORMATTER_OPTIONS.indent_continuation_lines),
        max_blank_lines:           getInt (sec, 'max_blank_lines',           DEFAULT_FORMATTER_OPTIONS.max_blank_lines),
        normalize_operator_spacing:getBool(sec, 'normalize_operator_spacing',DEFAULT_FORMATTER_OPTIONS.normalize_operator_spacing),
        normalize_comma_spacing:   getBool(sec, 'normalize_comma_spacing',   DEFAULT_FORMATTER_OPTIONS.normalize_comma_spacing),
        normalize_comment_space:   getBool(sec, 'normalize_comment_space',   DEFAULT_FORMATTER_OPTIONS.normalize_comment_space),
        expand_type_suffixes:      getBool(sec, 'expand_type_suffixes',      DEFAULT_FORMATTER_OPTIONS.expand_type_suffixes),
        format_on_save:            getBool(sec, 'format_on_save',            DEFAULT_FORMATTER_OPTIONS.format_on_save),
        format_on_pull:            sec['format_on_pull'] === undefined ? undefined : getBool(sec, 'format_on_pull', false),
    };
}

export type BreadcrumbsSortOrder = 'position' | 'name' | 'type';

export const DEFAULT_BREADCRUMBS_SORT: BreadcrumbsSortOrder = 'name';

const BREADCRUMBS_SORT_VALUES: readonly BreadcrumbsSortOrder[] = ['position', 'name', 'type'];

/** cat5dev.toml の [editor] breadcrumbs_sort を返す。
 *  toml・セクション・キーが無い場合や不正値の場合は undefined（＝何もしない）。 */
export function readBreadcrumbsSort(workspaceRoot: string): BreadcrumbsSortOrder | undefined {
    const tomlPath = path.join(workspaceRoot, 'cat5dev.toml');
    if (!fs.existsSync(tomlPath)) { return undefined; }

    let content: string;
    try { content = fs.readFileSync(tomlPath, 'utf-8'); } catch { return undefined; }

    const raw = parseToml(content)['editor']?.['breadcrumbs_sort'];
    if (raw === undefined) { return undefined; }
    const value = raw.replace(/^"|"$/g, '').trim().toLowerCase();
    return (BREADCRUMBS_SORT_VALUES as readonly string[]).includes(value)
        ? value as BreadcrumbsSortOrder
        : undefined;
}

export function gitignoreTemplate(): string {
    return `# OS
Thumbs.db

# Editor
.vscode/

# Logs
*.log
logs/

# Cache / temp
.cache/
tmp/
temp/

# Archives
*.zip
*.tar
*.gz

# Generated files
*.tmp
*.bak

# Claude Code
.claude/

# CAT5Dev settings
cat5dev.toml
`;
}

/**
 * cat5dev.toml の [project] セクションから設定を読み込む。
 * 旧版の `encoding` は使わない（Push / Pull は本文を Unicode のまま運ぶ）。残っていてもエラーにせず無視する。
 */
export function readProjectSettings(workspaceRoot: string): { targetProject: string; language: string } {
    const tomlPath = path.join(workspaceRoot, 'cat5dev.toml');
    if (!fs.existsSync(tomlPath)) { return { targetProject: '', language: '' }; }

    let content: string;
    try { content = fs.readFileSync(tomlPath, 'utf-8'); } catch { return { targetProject: '', language: '' }; }

    const proj = parseToml(content)['project'] ?? {};
    const targetProject = (proj['target_project'] ?? '').replace(/^"|"$/g, '');
    const language = (proj['language'] ?? '').replace(/^"|"$/g, '');
    return { targetProject, language };
}

/**
 * cat5dev.toml の [section] の key を `key = rawValue` に書き換える（コメント・他行を保持）。
 * rawValue は TOML の値をそのまま書く（文字列なら呼び出し側で "..." を付ける）。
 * キーが無ければセクション見出しの直後に、セクションも無ければ末尾に追加する。toml が無ければ何もしない。
 */
export function writeTomlKey(workspaceRoot: string, section: string, key: string, rawValue: string): void {
    const tomlPath = path.join(workspaceRoot, 'cat5dev.toml');
    if (!fs.existsSync(tomlPath)) { return; }

    const newLine = `${key} = ${rawValue}`;
    const lines = fs.readFileSync(tomlPath, 'utf-8').split(/\r?\n/);
    const sectionOf = (line: string): string | undefined => line.trim().match(/^\[([^\]]+)\]$/)?.[1].trim();

    let inSection = false;
    let written = false;
    let result = lines.map(line => {
        const sec = sectionOf(line);
        if (sec !== undefined) { inSection = sec === section; }
        if (inSection && !written) {
            const eqIdx = line.indexOf('=');
            if (eqIdx >= 0 && line.substring(0, eqIdx).trim() === key) {
                written = true;
                return newLine;
            }
        }
        return line;
    });

    if (!written) {
        const headerIdx = result.findIndex(line => sectionOf(line) === section);
        if (headerIdx >= 0) {
            result.splice(headerIdx + 1, 0, newLine);
        } else {
            // 末尾の空行の後ろにまとめて付ける
            while (result.length > 0 && result[result.length - 1].trim() === '') { result.pop(); }
            result = [...result, '', `[${section}]`, newLine, ''];
        }
    }

    fs.writeFileSync(tomlPath, result.join('\n'), 'utf-8');
}

/** cat5dev.toml の [project] セクションの指定キーの値（文字列）を書き換える（コメント・他行を保持） */
export function writeTomlProjectKey(workspaceRoot: string, key: string, value: string): void {
    writeTomlKey(workspaceRoot, 'project', key, `"${value}"`);
}

const tomlComments = {
    ja: {
        header: '# Cat5Dev 設定ファイル',
        targetProject: '# 対象の CATIA VBA プロジェクト名',
        language: '# 言語設定 (ja / en)',
        optionExplicit: '# Option Explicit が宣言されていない場合に警告',
        onErrorResumeNext: '# On Error Resume Next の使用時に警告',
        goto: '# GoTo の使用時に警告 (On Error GoTo は除外)',
        maxLineLength: '# 指定の文字数を超える行がある場合に警告 (0 = 無効)',
        unusedVariables: '# Dim で宣言された変数が使用されていない場合に警告',
        maxNestingDepth: '# ネスト深さが閾値を超えた場合に警告 (0 = 無効)',
        maxFunctionLines: '# Sub/Function の行数が閾値を超えた場合に警告 (0 = 無効)',
        unmatchedParens: '# 括弧の不一致をエラーとして報告',
        unmatchedBlocks: '# End If / End Sub / End Function 等の欠落をエラーとして報告',
        nonAnsiCharacter: '# CATIA の VBA（Windows の ANSI コードページ）で表せない文字を警告 (Push すると ? などに置き換わるため)',
        ansiCodePage: '# non_ansi_character の判定に使うコードページ (932 = 日本語、1252 = 英語、936 = 簡体字中国語、0 = システムの設定を自動検出 など)\n# 言語を切り替えると、その言語の値 (ja = 932 / en = 1252 / zh = 936) に書き換わります\n# [project] の旧 encoding とは別の設定です',
        formatterEnabled: '# フォーマッターの有効/無効',
        indentSize: '# インデントのスペース数',
        capitalizeKeywords: '# VBA キーワードを大文字化 (If, Dim, Sub 等)',
        fixIndentation: '# インデントの修正',
        trimTrailingSpace: '# 行末の空白を削除',
        ensureContinuationSpace: '# 継続文字 (_) の前にスペースを入れる',
        indentContinuationLines: '# 継続行を1レベルインデント',
        maxBlankLines: '# 連続する空白行の最大数 (0 = 無効)',
        normalizeOperatorSpacing: '# 演算子 (=, +, - 等) 周囲のスペースを正規化',
        normalizeCommaSpacing: '# カンマ後のスペースを正規化',
        normalizeCommentSpace: '# コメント文字 (\') の後にスペースを入れる',
        expandTypeSuffixes: '# 型サフィックスの短縮形を展開 (% → Integer, $ → String 等)',
        formatOnSave: '# 保存時に自動フォーマット',
        formatOnPull: '# CATIA から Pull したファイルを自動フォーマット',
        breadcrumbsSort: '# パンくずリストのシンボル並び順 (position = 記述順 / name = 名前順 / type = 種類順)\n# ワークスペース設定 (.vscode/settings.json) に反映され、このワークスペースの全ファイルに適用されます',
    },
    en: {
        header: '# Cat5Dev configuration file',
        targetProject: '# Target CATIA VBA project name',
        language: '# Language (ja / en)',
        optionExplicit: '# Warn when Option Explicit is not declared',
        onErrorResumeNext: '# Warn on usage of On Error Resume Next',
        goto: '# Warn on GoTo usage (On Error GoTo is excluded)',
        maxLineLength: '# Warn when a line exceeds the specified character count (0 = disabled)',
        unusedVariables: '# Warn when a variable is declared with Dim but never used',
        maxNestingDepth: '# Warn when nesting depth exceeds the threshold (0 = disabled)',
        maxFunctionLines: '# Warn when a Sub/Function exceeds the line count threshold (0 = disabled)',
        unmatchedParens: '# Report mismatched parentheses as an error',
        unmatchedBlocks: '# Report missing End If / End Sub / End Function etc. as an error',
        nonAnsiCharacter: '# Warn about characters that CATIA VBA (the Windows ANSI code page) cannot represent (Push replaces them with ? or similar)',
        ansiCodePage: '# Code page used by non_ansi_character (932 = Japanese, 1252 = English, 936 = Simplified Chinese, 0 = detect the system setting, etc.)\n# Switching the language rewrites this to the value for that language (ja = 932 / en = 1252 / zh = 936)\n# Not related to the former [project] encoding setting',
        formatterEnabled: '# Enable or disable the formatter entirely',
        indentSize: '# Number of spaces per indentation level',
        capitalizeKeywords: '# Capitalize VBA keywords (If, Dim, Sub, etc.)',
        fixIndentation: '# Fix incorrect indentation',
        trimTrailingSpace: '# Remove trailing whitespace from each line',
        ensureContinuationSpace: '# Ensure a space before line continuation character (_)',
        indentContinuationLines: '# Indent continuation lines by one level',
        maxBlankLines: '# Maximum number of consecutive blank lines (0 = disabled)',
        normalizeOperatorSpacing: '# Normalize spacing around operators (=, +, -, etc.)',
        normalizeCommaSpacing: '# Normalize spacing after commas',
        normalizeCommentSpace: '# Ensure a space after comment character (\')',
        expandTypeSuffixes: '# Expand type suffix shorthand (% → Integer, $ → String, etc.)',
        formatOnSave: '# Automatically format on save',
        formatOnPull: '# Automatically format files pulled from CATIA',
        breadcrumbsSort: '# Symbol order in the breadcrumbs (position = source order / name = by name / type = by kind)\n# Written to the workspace settings (.vscode/settings.json) and applies to every file in this workspace',
    },
    zh: {
        header: '# Cat5Dev 配置文件',
        targetProject: '# 目标 CATIA VBA 项目名称',
        language: '# 语言设置 (ja / en / zh)',
        optionExplicit: '# 当未声明 Option Explicit 时发出警告',
        onErrorResumeNext: '# 使用 On Error Resume Next 时发出警告',
        goto: '# 使用 GoTo 时发出警告 (不包括 On Error GoTo)',
        maxLineLength: '# 单行字符数超过指定值时发出警告 (0 = 禁用)',
        unusedVariables: '# 变量已声明但未使用时发出警告',
        maxNestingDepth: '# 嵌套深度超过阈值时发出警告 (0 = 禁用)',
        maxFunctionLines: '# Sub/Function 行数超过阈值时发出警告 (0 = 禁用)',
        unmatchedParens: '# 将括号不匹配报告为错误',
        unmatchedBlocks: '# 将缺少 End If / End Sub / End Function 等报告为错误',
        nonAnsiCharacter: '# 对 CATIA 的 VBA（Windows 的 ANSI 代码页）无法表示的字符发出警告 (Push 后会被替换为 ? 等字符)',
        ansiCodePage: '# non_ansi_character 判断所用的代码页 (936 = 简体中文，932 = 日语，1252 = 英语，0 = 自动检测系统设置 等)\n# 切换语言时会改写为该语言对应的值 (ja = 932 / en = 1252 / zh = 936)\n# 与旧版 [project] 的 encoding 设置无关',
        formatterEnabled: '# 启用或禁用格式化程序',
        indentSize: '# 每个缩进级别的空格数',
        capitalizeKeywords: '# 自动大写 VBA 关键字 (If, Dim, Sub 等)',
        fixIndentation: '# 自动修正缩进',
        trimTrailingSpace: '# 删除行尾空白字符',
        ensureContinuationSpace: '# 确保换行符 (_) 前有空格',
        indentContinuationLines: '# 将换行行缩进一个级别',
        maxBlankLines: '# 连续空行的最大数量 (0 = 禁用)',
        normalizeOperatorSpacing: '# 规范运算符 (=, +, - 等) 两侧的空格',
        normalizeCommaSpacing: '# 规范逗号后的空格',
        normalizeCommentSpace: '# 确保注释符号 (\') 后有空格',
        expandTypeSuffixes: '# 展开类型后缀简写 (% → Integer, $ → String 等)',
        formatOnSave: '# 保存时自动格式化',
        formatOnPull: '# 自动格式化从 CATIA 拉取的文件',
        breadcrumbsSort: '# 面包屑导航中符号的排序方式 (position = 代码顺序 / name = 按名称 / type = 按类型)\n# 写入工作区设置 (.vscode/settings.json)，并应用于此工作区中的所有文件',
    },
};

/** cat5dev.toml の雛形テキストを返す。値は DEFAULT_*_OPTIONS から生成する */
export function tomlTemplate(lang: Language = 'ja'): string {
    const l = DEFAULT_LINT_OPTIONS;
    const f = DEFAULT_FORMATTER_OPTIONS;
    const c = tomlComments[lang];
    return `${c.header}

[project]
${c.targetProject}
target_project = ""

${c.language}
language = "${lang}"

[lint]
enabled = false

[lint.rules]
${c.optionExplicit}
option_explicit = ${l.option_explicit}

${c.onErrorResumeNext}
on_error_resume_next = ${l.on_error_resume_next}

${c.goto}
goto = ${l.goto}

${c.maxLineLength}
max_line_length = ${l.max_line_length}

${c.unusedVariables}
unused_variables = ${l.unused_variables}

${c.maxNestingDepth}
max_nesting_depth = ${l.max_nesting_depth}

${c.maxFunctionLines}
max_function_lines = ${l.max_function_lines}

${c.unmatchedParens}
unmatched_parens = ${l.unmatched_parens}

${c.unmatchedBlocks}
unmatched_blocks = ${l.unmatched_blocks}

${c.nonAnsiCharacter}
non_ansi_character = ${l.non_ansi_character}

${c.ansiCodePage}
ansi_code_page = ${codePageForLanguage(lang)}

[formatter]
${c.formatterEnabled}
enabled = false

${c.indentSize}
indent_size = ${f.indent_size}

${c.capitalizeKeywords}
capitalize_keywords = ${f.capitalize_keywords}

${c.fixIndentation}
fix_indentation = ${f.fix_indentation}

${c.trimTrailingSpace}
trim_trailing_space = ${f.trim_trailing_space}

${c.ensureContinuationSpace}
ensure_continuation_space = ${f.ensure_continuation_space}

${c.indentContinuationLines}
indent_continuation_lines = ${f.indent_continuation_lines}

${c.maxBlankLines}
max_blank_lines = ${f.max_blank_lines}

${c.normalizeOperatorSpacing}
normalize_operator_spacing = ${f.normalize_operator_spacing}

${c.normalizeCommaSpacing}
normalize_comma_spacing = ${f.normalize_comma_spacing}

${c.normalizeCommentSpace}
normalize_comment_space = ${f.normalize_comment_space}

${c.expandTypeSuffixes}
expand_type_suffixes = ${f.expand_type_suffixes}

${c.formatOnSave}
format_on_save = ${f.format_on_save}

${c.formatOnPull}
format_on_pull = ${f.format_on_pull}

[editor]
${c.breadcrumbsSort}
breadcrumbs_sort = "${DEFAULT_BREADCRUMBS_SORT}"
`;
}
