import { execFile } from 'child_process';
import * as iconv from 'iconv-lite';

/**
 * VBA はモジュール本文をシステムの ANSI コードページで保持する。
 * そのコードページで往復できない文字は、Push した時点で ? などの別の文字に置き換わる。
 * このモジュールは、その判定（Lint ルール VBA010 non_ansi_character）と、
 * システムの ANSI コードページの検出をまとめる。vscode には依存しない（単体テストのため）。
 */

/** Lint ルール non_ansi_character の診断コード */
export const NON_ANSI_DIAGNOSTIC_CODE = 'VBA010';

/** 対応する ANSI コードページと iconv-lite の符号化名 */
const ICONV_NAMES: Readonly<Record<number, string>> = {
    874: 'cp874',    // タイ語
    932: 'cp932',    // 日本語（Shift-JIS）
    936: 'cp936',    // 簡体字中国語（GBK）
    949: 'cp949',    // 韓国語
    950: 'cp950',    // 繁体字中国語（Big5）
    1250: 'cp1250',  // 中欧
    1251: 'cp1251',  // キリル
    1252: 'cp1252',  // 西欧
    1253: 'cp1253',  // ギリシャ語
    1254: 'cp1254',  // トルコ語
    1255: 'cp1255',  // ヘブライ語
    1256: 'cp1256',  // アラビア語
    1257: 'cp1257',  // バルト
    1258: 'cp1258',  // ベトナム語
};

/** 対応しているコードページの一覧（文書・ログ用） */
export const SUPPORTED_CODE_PAGES: readonly number[] = Object.keys(ICONV_NAMES).map(Number);

/** コードページに対応する iconv-lite の符号化名。対応外なら undefined */
export function iconvNameForCodePage(codePage: number): string | undefined {
    const name = ICONV_NAMES[codePage];
    return name !== undefined && iconv.encodingExists(name) ? name : undefined;
}

/** 1 文字ずつ「コードページで往復できるか」を判定する。結果はコードポイントごとにキャッシュする */
export class AnsiCharChecker {
    private readonly cache = new Map<number, boolean>();

    private constructor(readonly codePage: number, private readonly encoding: string) { }

    /** 対応外のコードページなら undefined */
    static create(codePage: number): AnsiCharChecker | undefined {
        const enc = iconvNameForCodePage(codePage);
        return enc === undefined ? undefined : new AnsiCharChecker(codePage, enc);
    }

    /** ch（1 コードポイント分の文字列）を encode → decode して元に戻るか */
    canRoundTrip(ch: string): boolean {
        const cp = ch.codePointAt(0) ?? 0;
        if (cp < 0x80) { return true; }
        const cached = this.cache.get(cp);
        if (cached !== undefined) { return cached; }
        // 対になっていないサロゲートは符号化できない
        const ok = !(cp >= 0xD800 && cp <= 0xDFFF)
            && iconv.decode(iconv.encode(ch, this.encoding), this.encoding) === ch;
        this.cache.set(cp, ok);
        return ok;
    }
}

const checkers = new Map<number, AnsiCharChecker | undefined>();

/** コードページごとの判定器（キャッシュ込み）を共有して返す。対応外なら undefined */
export function getAnsiCharChecker(codePage: number): AnsiCharChecker | undefined {
    if (!checkers.has(codePage)) {
        checkers.set(codePage, AnsiCharChecker.create(codePage));
    }
    return checkers.get(codePage);
}

/** 往復できない文字の範囲（行・列は 0 始まり、列は UTF-16 単位で VSCode の Position と同じ） */
export interface NonAnsiRange {
    line: number;
    startCol: number;
    endCol: number;
    /** 範囲内の文字そのもの */
    text: string;
}

const NON_ASCII_RUN = /[^\x00-\x7F]+/g;

/**
 * text の中で、コードページで往復できない文字の範囲を返す。
 * 連続する文字は 1 つの範囲にまとめる。ASCII だけの行はすぐに飛ばす。
 */
export function findNonAnsiRanges(text: string, checker: AnsiCharChecker): NonAnsiRange[] {
    const result: NonAnsiRange[] = [];
    const lines = text.split(/\r\n|\r|\n/);
    for (let line = 0; line < lines.length; line++) {
        const s = lines[line];
        NON_ASCII_RUN.lastIndex = 0;
        let m: RegExpExecArray | null;
        while ((m = NON_ASCII_RUN.exec(s)) !== null) {
            const runStart = m.index;
            const runEnd = runStart + m[0].length;
            let badStart = -1;
            let i = runStart;
            while (i < runEnd) {
                const cp = s.codePointAt(i)!;
                const width = cp > 0xFFFF ? 2 : 1;
                const ok = checker.canRoundTrip(s.slice(i, i + width));
                if (!ok && badStart < 0) {
                    badStart = i;
                } else if (ok && badStart >= 0) {
                    result.push({ line, startCol: badStart, endCol: i, text: s.substring(badStart, i) });
                    badStart = -1;
                }
                i += width;
            }
            if (badStart >= 0) {
                result.push({ line, startCol: badStart, endCol: runEnd, text: s.substring(badStart, runEnd) });
            }
        }
    }
    return result;
}

/** システムの ANSI コードページを検出できなかったときに、表示言語（VSCode の env.language など）から推定する */
export function estimateCodePageFromLanguage(language: string | undefined): number {
    const lang = (language ?? '').toLowerCase();
    if (lang.startsWith('zh')) {
        return /^zh-(tw|hk|mo)|hant/.test(lang) ? 950 : 936;
    }
    if (lang.startsWith('ko')) { return 949; }
    // 日本語、その他・不明は 932（この拡張機能の主な利用環境）
    return 932;
}

/** reg.exe の出力から ACP の値を取り出す */
export function parseAcpFromRegOutput(output: string): number | undefined {
    const m = /^\s*ACP\s+REG_\w+\s+(\d+)\s*$/im.exec(output);
    if (!m) { return undefined; }
    const n = parseInt(m[1], 10);
    return n > 0 ? n : undefined;
}

export type CodePageSource = 'registry' | 'fallback';

export interface DetectedCodePage {
    codePage: number;
    source: CodePageSource;
    /** fallback になった理由（ログ用） */
    reason?: string;
}

export type RegQuery = () => Promise<string>;

const ACP_KEY = 'HKLM\\SYSTEM\\CurrentControlSet\\Control\\Nls\\CodePage';

/** reg.exe でシステムの ANSI コードページ（ACP）を読む */
export const queryAcpWithRegExe: RegQuery = () => new Promise((resolve, reject) => {
    execFile('reg.exe', ['query', ACP_KEY, '/v', 'ACP'], { timeout: 5000, windowsHide: true }, (err, stdout) => {
        if (err) { reject(err); return; }
        resolve(String(stdout));
    });
});

/**
 * システムの ANSI コードページを検出する。失敗したら表示言語から推定する（失敗しても reject しない）。
 * 拡張機能の起動時に 1 回だけ呼ぶ想定。
 */
export async function detectSystemAnsiCodePage(
    fallbackLanguage: string | undefined,
    query: RegQuery = queryAcpWithRegExe,
): Promise<DetectedCodePage> {
    const fallback = (reason: string): DetectedCodePage =>
        ({ codePage: estimateCodePageFromLanguage(fallbackLanguage), source: 'fallback', reason });
    if (query === queryAcpWithRegExe && process.platform !== 'win32') {
        return fallback(`platform ${process.platform}`);
    }
    try {
        const cp = parseAcpFromRegOutput(await query());
        return cp === undefined ? fallback('ACP not found') : { codePage: cp, source: 'registry' };
    } catch (err) {
        return fallback(err instanceof Error ? err.message : String(err));
    }
}
