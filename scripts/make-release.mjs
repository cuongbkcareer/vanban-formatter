// Đóng gói file zip phát hành vào dist/. Chạy sau `npm run build` (hoặc dùng `npm run release`).
import JSZip from 'jszip';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { DEFAULT_BASE_URL, readVersion } from './config.mjs';

const version = await readVersion();
const root = `vanban-formatter-addin-v${version}`;
const zip = new JSZip();

const read = (p) => readFile(p);
const BOM = Buffer.from([0xef, 0xbb, 0xbf]);

// manifest + hướng dẫn + script cài đặt
zip.file(`${root}/manifest.xml`, await read('docs/manifest.xml'));
zip.file(`${root}/HUONG-DAN-CAI-DAT.md`, await read('HUONG-DAN-CAI-DAT.md'));
zip.file(`${root}/CAI-DAT-MAC.command`, await read('installer/CAI-DAT-MAC.command'), { unixPermissions: 0o755 });
// PowerShell 5.1 đọc file không có BOM theo bảng mã ANSI nên làm hỏng tiếng Việt.
zip.file(`${root}/CAI-DAT-WINDOWS.ps1`, Buffer.concat([BOM, await read('installer/CAI-DAT-WINDOWS.ps1')]));
zip.file(`${root}/CAI-DAT-WINDOWS.bat`, (await read('installer/CAI-DAT-WINDOWS.bat')).toString('utf8').replace(/\r?\n/g, '\r\n'));

// bản web độc lập (chạy offline) và bản để tự host
zip.file(`${root}/vanban.html`, await read('docs/vanban.html'));
for (const f of ['taskpane.html', 'taskpane.js', 'assets/icon-16.png', 'assets/icon-32.png', 'assets/icon-80.png']) {
  zip.file(`${root}/tu-host/${f}`, await read(`docs/${f}`));
}

const data = await zip.generateAsync({
  type: 'nodebuffer',
  platform: 'UNIX',
  compression: 'DEFLATE',
  compressionOptions: { level: 9 },
});
await mkdir('dist', { recursive: true });
const out = `dist/${root}.zip`;
await writeFile(out, data);
await writeFile('dist/manifest.xml', await read('docs/manifest.xml'));
console.log(`${out} (${(data.length / 1024).toFixed(0)} KB), manifest trỏ tới ${DEFAULT_BASE_URL}`);
