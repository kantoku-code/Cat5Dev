import * as path from 'path';

export type CatiaStep = 'list' | 'pull' | 'check' | 'push';
export interface CatiaItem { name: string; type: 1 | 2 | 3; file: string }
export interface CatiaFailure { phase: string; target: string; code: number; message: string }
export interface CatiaExpectation {
    operationId: string; stepId: string; kind: string; targetProject: string;
}
/** A module sent by Push: the local text as-is (UTF-16 in request.json; no encoding conversion). */
export interface CatiaModuleText { name: string; type: 1 | 2 | 3; text: string }
/** One catia-run.exe step. directory holds request.json, result.json and the exported m<i>.txt (UTF-8). */
export interface CatiaRequest extends CatiaExpectation {
    kind: CatiaStep;
    directory: string;
    modules?: CatiaModuleText[];
    deletions?: string[];
}
/** VBA project name of the Cat5Dev bridge (C5D_Bridge.catvba); never offered as a sync target. */
export const BRIDGE_PROJECT_NAME = 'C5D_Bridge';
export interface CatiaResult extends CatiaExpectation {
    version: 1;
    expectedCount: number;
    succeededCount: number;
    failures: CatiaFailure[];
    items: CatiaItem[];
    projects: string[];
    updatedCount: number;
    deletedCount: number;
}

export function safeFileName(name: string): boolean {
    return !!name && name !== '.' && name !== '..' && path.basename(name) === name
        && !/[\\/:\x00-\x1f<>"|?*]/.test(name) && !/[. ]$/.test(name)
        && !/^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(name);
}

/** Comparison form of module source: line endings unified to LF and trailing whitespace removed. */
export function normalizeModuleText(text: string): string {
    return text.replace(/\r\n?/g, '\n').trimEnd();
}

/** Removes a leading BOM (U+FEFF) that editors may add when saving as UTF-8 with BOM. */
export function stripBom(text: string): string {
    return text.charCodeAt(0) === 0xFEFF ? text.slice(1) : text;
}

/** Validates terminal results, including confirmed partial failures. */
export function validateCatiaResult(text: string, expected: CatiaExpectation): CatiaResult {
    const r = JSON.parse(stripBom(text));
    const bad = (): never => { throw new Error('Invalid CATIA execution result'); };
    if (!r || typeof r !== 'object' || r.version !== 1) { return bad(); }
    for (const key of ['operationId', 'stepId', 'kind', 'targetProject'] as const) {
        if (r[key] !== expected[key]) { return bad(); }
    }
    for (const key of ['expectedCount', 'succeededCount', 'updatedCount', 'deletedCount']) {
        if (!Number.isSafeInteger(r[key]) || r[key] < 0) { return bad(); }
    }
    if (r.succeededCount > r.expectedCount || !Array.isArray(r.failures)
        || !Array.isArray(r.items) || !Array.isArray(r.projects)) { return bad(); }
    if (r.failures.some((f: any) => !f || typeof f.phase !== 'string'
        || typeof f.target !== 'string' || !Number.isSafeInteger(f.code)
        || typeof f.message !== 'string')) { return bad(); }
    if (!r.failures.length && r.expectedCount !== r.succeededCount) { return bad(); }
    if (r.projects.some((p: any) => typeof p !== 'string' || !p)) { return bad(); }
    const names = new Set<string>();
    const files = new Set<string>();
    for (const item of r.items) {
        if (!item || typeof item.name !== 'string' || !safeFileName(item.name)
            || ![1, 2, 3].includes(item.type) || typeof item.file !== 'string'
            || !safeFileName(item.file)) { return bad(); }
        const name = item.name.toLowerCase();
        const file = item.file.toLowerCase();
        if (names.has(name) || files.has(file)) { return bad(); }
        names.add(name); files.add(file);
    }
    if (expected.kind === 'list' && r.projects.length !== r.succeededCount) { return bad(); }
    if (['pull', 'check'].includes(expected.kind) && r.items.length !== r.succeededCount) { return bad(); }
    if (expected.kind === 'push' && r.updatedCount + r.deletedCount !== r.succeededCount) { return bad(); }
    return r as CatiaResult;
}

export function catiaSucceeded(result: CatiaResult): boolean {
    return result.failures.length === 0 && result.expectedCount === result.succeededCount;
}
