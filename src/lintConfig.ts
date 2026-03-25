import * as fs from 'fs';
import * as path from 'path';

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
}

export const DEFAULT_LINT_OPTIONS: LintOptions = {
    option_explicit: true,
    on_error_resume_next: true,
    goto: true,
    max_line_length: 200,
    unused_variables: true,
    max_nesting_depth: 5,
    max_function_lines: 100,
    unmatched_parens: true,
    unmatched_blocks: true,
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

/** ワークスペースルートの cat5dev.toml を読み込んで LintOptions を返す */
export function readLintOptions(workspaceRoot: string): LintOptions {
    const tomlPath = path.join(workspaceRoot, 'cat5dev.toml');
    if (!fs.existsSync(tomlPath)) {
        return { ...DEFAULT_LINT_OPTIONS };
    }

    let content: string;
    try {
        content = fs.readFileSync(tomlPath, 'utf-8');
    } catch {
        return { ...DEFAULT_LINT_OPTIONS };
    }

    const toml = parseToml(content);

    // [lint] セクションで enabled = false なら全ルール無効
    const lintSection = toml['lint'] ?? {};
    if (getBool(lintSection, 'enabled', true) === false) {
        return {
            option_explicit: false,
            on_error_resume_next: false,
            goto: false,
            max_line_length: 0,
            unused_variables: false,
            max_nesting_depth: 0,
            max_function_lines: 0,
            unmatched_parens: false,
            unmatched_blocks: false,
        };
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
    };
}

/** cat5dev.toml の雛形テキストを返す */
export function tomlTemplate(): string {
    return `# Cat5Dev configuration file

[lint]
enabled = true

[lint.rules]
# Warn when Option Explicit is not declared
option_explicit = true

# Warn on usage of On Error Resume Next
on_error_resume_next = true

# Warn on GoTo usage (On Error GoTo is excluded)
goto = true

# Warn when a line exceeds the specified character count (0 = disabled)
max_line_length = 200

# Warn when a variable is declared with Dim but never used
unused_variables = true

# Warn when nesting depth exceeds the threshold (0 = disabled)
max_nesting_depth = 5

# Warn when a Sub/Function exceeds the line count threshold (0 = disabled)
max_function_lines = 100

# Report mismatched parentheses as an error
unmatched_parens = true

# Report missing End If / End Sub / End Function etc. as an error
unmatched_blocks = true
`;
}
