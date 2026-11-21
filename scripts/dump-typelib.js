// Type library dump (developer tool, not shipped): lists every interface/coclass/enum member (name, DISPID, arguments, flags) of the
// VBA 7.1 / APC type libraries, to look for a way to close or remove a VBA project (C5D_Bridge) without quitting CATIA.
// Usage: node scripts/dump-typelib.js [<file>[\<resource index>] ...]
//   default files: apc71.dll, VBE7.DLL, VBEUI.DLL and 1033\*.DLL of C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1,
//   plus VBE6EXT.OLB (VBIDE 5.3, 32-bit VBA6) when it exists. Every TYPELIB resource 1..16 of each file is tried.
// Builds tools/catia-run/TypeLibDump.cs into temp/typelib-dump.exe (.NET Framework 4.x csc.exe) and writes its output to
// temp/typelib-apc71-<yyyyMMdd-HHmmss>.txt. LoadTypeLibEx(REGKIND_NONE) only: nothing is registered, no COM object is created,
// CATIA is never contacted.
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { timestamp } = require('./probe-catscript');

const ROOT = path.resolve(__dirname, '..');
const VBA71 = 'C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1';
const VBE6EXT = 'C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB';

/** Default input files (only the existing ones). */
function defaultFiles() {
    const files = ['apc71.dll', 'VBE7.DLL', 'VBEUI.DLL'].map(name => path.join(VBA71, name));
    const intl = path.join(VBA71, '1033');
    if (fs.existsSync(intl)) files.push(...fs.readdirSync(intl).filter(name => /\.dll$/i.test(name)).map(name => path.join(intl, name)));
    files.push(VBE6EXT);
    return files.filter(file => fs.existsSync(file));
}

function buildDumper(root = ROOT) {
    const csc = path.join(process.env.SystemRoot || 'C:\\Windows', 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe');
    if (!fs.existsSync(csc)) throw new Error(`csc.exe not found: ${csc}`);
    const exe = path.join(root, 'temp', 'typelib-dump.exe');
    fs.mkdirSync(path.dirname(exe), { recursive: true });
    try {
        execFileSync(csc, ['/nologo', '/optimize+', '/target:exe', '/platform:anycpu', `/out:${exe}`,
            path.join(root, 'tools', 'catia-run', 'TypeLibDump.cs')], { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
    } catch (error) {
        throw new Error(`csc failed: ${String(error.stdout || '') + String(error.stderr || '') || error}`);
    }
    return exe;
}

function main() {
    const files = process.argv.slice(2).length ? process.argv.slice(2) : defaultFiles();
    if (!files.length) { console.error(`no type library found under ${VBA71}`); process.exitCode = 1; return; }
    const exe = buildDumper();
    let stdout;
    try {
        stdout = execFileSync(exe, files, { windowsHide: true, encoding: 'utf8', maxBuffer: 256 * 1024 * 1024 });
    } catch (error) {
        stdout = String(error.stdout || '');
        process.exitCode = 1;
    }
    const out = path.join(ROOT, 'temp', `typelib-apc71-${timestamp()}.txt`);
    fs.writeFileSync(out, `# node scripts/dump-typelib.js  ${new Date().toISOString()}\n# files: ${files.join(' | ')}\n${stdout}`);
    const highlight = stdout.slice(stdout.lastIndexOf('\n== highlight'));
    console.log(`written: ${out} (${stdout.length} chars)`);
    console.log(stdout.split('\n').filter(line => line.startsWith('== ')).join('\n'));
    console.log(highlight.split('\n').slice(0, 400).join('\n'));
}
if (require.main === module) main();
module.exports = { defaultFiles, buildDumper };
