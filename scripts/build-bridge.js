// Builds bin/catia-run.exe with the csc.exe that ships with .NET Framework 4.x (no SDK required, C# 5).
// Sources: tools/catia-run/CatiaRun.cs and tools/catia-run/CatiaRun.*.cs (EvalProbe.cs / TypeLibDump.cs are separate tools).
// The runtime module resources/bridge/C5D_Bridge_Runtime.bas is embedded as the manifest resource "C5D_Bridge_Runtime.bas"
// (docs/dev/bridge-contract.md, chapter 3).
// Usage: node scripts/build-bridge.js [--out <path.exe>]   (default bin/catia-run.exe; --out is for tests)
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');

const root = path.resolve(__dirname, '..');
const sourceDir = path.join(root, 'tools', 'catia-run');
const runtimeBas = path.join(root, 'resources', 'bridge', 'C5D_Bridge_Runtime.bas');
const RESOURCE_NAME = 'C5D_Bridge_Runtime.bas';

/** csc.exe of .NET Framework 4.x (64-bit first). */
function findCsc(env = process.env) {
    const windows = env.SystemRoot || 'C:\\Windows';
    for (const framework of ['Framework64', 'Framework']) {
        const csc = path.join(windows, 'Microsoft.NET', framework, 'v4.0.30319', 'csc.exe');
        if (fs.existsSync(csc)) return csc;
    }
    return null;
}

/** CatiaRun.cs first, then CatiaRun.*.cs in name order. */
function sourceFiles(dir = sourceDir) {
    const names = fs.readdirSync(dir).filter(name => /^CatiaRun(?:\.[^.]+)*\.cs$/i.test(name) && name !== 'CatiaRun.cs').sort();
    return [path.join(dir, 'CatiaRun.cs'), ...names.map(name => path.join(dir, name))];
}

/** csc.exe arguments for out (an absolute .exe path). */
function cscArguments(out, sources = sourceFiles()) {
    return ['/nologo', '/codepage:65001', '/optimize+', '/warnaserror+', '/target:exe', '/platform:anycpu',
        `/resource:${runtimeBas},${RESOURCE_NAME}`, `/out:${out}`, ...sources];
}

function main(argv) {
    let out = path.join(root, 'bin', 'catia-run.exe');
    for (let i = 0; i < argv.length; i++) {
        if (argv[i] === '--out' && i + 1 < argv.length) { out = path.resolve(argv[++i]); continue; }
        console.error(`unknown argument: ${argv[i]}\nusage: node scripts/build-bridge.js [--out <path.exe>]`);
        return 2;
    }
    const csc = findCsc();
    if (!csc) { console.error('csc.exe of .NET Framework 4.x not found'); return 1; }
    if (!fs.existsSync(runtimeBas)) { console.error(`runtime module not found: ${runtimeBas}`); return 1; }
    fs.mkdirSync(path.dirname(out), { recursive: true });
    try {
        execFileSync(csc, cscArguments(out), { stdio: 'inherit', windowsHide: true });
    } catch (e) {
        console.error(`csc.exe failed (exit ${e.status ?? e.code})`);
        return 1;
    }
    return 0;
}

if (require.main === module) process.exitCode = main(process.argv.slice(2));

module.exports = { findCsc, sourceFiles, cscArguments, RESOURCE_NAME };
