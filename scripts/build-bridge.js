// Builds bin/catia-run.exe with the csc.exe that ships with .NET Framework 4.x (no SDK required).
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const root = path.resolve(__dirname, '..');
const csc = path.join(process.env.SystemRoot || 'C:\\Windows', 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe');
if (!fs.existsSync(csc)) { console.error(`csc.exe not found: ${csc}`); process.exit(1); }
fs.mkdirSync(path.join(root, 'bin'), { recursive: true });
execFileSync(csc, ['/nologo', '/optimize+', '/target:exe', '/platform:anycpu',
    `/out:${path.join(root, 'bin', 'catia-run.exe')}`, path.join(root, 'tools', 'catia-run', 'CatiaRun.cs')], { stdio: 'inherit' });
