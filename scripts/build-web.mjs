// Đóng gói giao diện vào docs/ (thư mục GitHub Pages phục vụ):
//   docs/vanban.html     bản web độc lập, một file, mở thẳng bằng trình duyệt (CSP chặn mọi kết nối mạng)
//   docs/taskpane.html   khung bên cạnh trong Word (+ taskpane.js, assets/, manifest.xml)
import { build } from 'esbuild';
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { DEFAULT_BASE_URL, REPO, readVersion } from './config.mjs';
import { generateManifest } from './manifest.mjs';

const OUT = 'docs';
const baseUrl = process.env.ADDIN_BASE_URL ?? DEFAULT_BASE_URL;

async function bundle(entry) {
  const r = await build({
    entryPoints: [entry],
    bundle: true,
    write: false,
    format: 'iife',
    platform: 'browser',
    target: 'es2020',
    minify: true,
    legalComments: 'none',
  });
  return r.outputFiles[0].text;
}

const template = await readFile('web/template.html', 'utf8');
const fill = (csp, script) =>
  template.replace('<!--CSP-->', () => csp).replace('<!--SCRIPT-->', () => script);

// Bản độc lập: không kết nối mạng.
const STANDALONE_CSP =
  `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data:; connect-src 'none'; form-action 'none'; base-uri 'none'">`;

// Khung bên cạnh Word: chỉ tải office.js từ CDN của Microsoft (bắt buộc), mã của chúng ta là file cùng nguồn.
// connect-src cho phép các tên miền Microsoft vì chính office.js cần liên lạc; mã của dự án không gọi mạng (có test kiểm tra).
const ADDIN_CSP =
  `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'self' https://appsforoffice.microsoft.com; style-src 'unsafe-inline'; img-src 'self' data:; connect-src 'self' https://*.microsoft.com https://*.office.com https://*.office365.com https://*.officeapps.live.com; form-action 'none'; base-uri 'none'">`;

const standaloneJs = (await bundle('web/standalone.ts')).replace(/<\/script/gi, '<\\/script');
const addinJs = await bundle('web/addin.ts');

await mkdir(`${OUT}/assets`, { recursive: true });
await writeFile(`${OUT}/vanban.html`, fill(STANDALONE_CSP, `<script>${standaloneJs}</script>`));
await writeFile(
  `${OUT}/taskpane.html`,
  fill(
    ADDIN_CSP,
    `<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>\n<script src="taskpane.js"></script>`,
  ),
);
await writeFile(`${OUT}/taskpane.js`, addinJs);
for (const s of [16, 32, 80]) await copyFile(`assets/icon-${s}.png`, `${OUT}/assets/icon-${s}.png`);
await writeFile(`${OUT}/manifest.xml`, generateManifest({ baseUrl, version: await readVersion() }));
await writeFile(`${OUT}/.nojekyll`, '');

// Trang chủ GitHub Pages: liên kết tải và hướng dẫn.
const version = await readVersion();
await writeFile(
  `${OUT}/index.html`,
  `<!DOCTYPE html>
<html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'">
<title>Chuẩn hoá văn bản hành chính (NĐ 30/2020)</title>
<style>body{font:16px/1.6 system-ui,-apple-system,"Segoe UI",sans-serif;max-width:720px;margin:40px auto;padding:0 16px}a.btn{display:inline-block;background:#1a56c4;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none;margin:4px 8px 4px 0}code{background:#8882;padding:1px 5px;border-radius:4px}</style></head>
<body>
<h1>Chuẩn hoá văn bản hành chính</h1>
<p>Chuyển văn bản thô thành văn bản hành chính đúng thể thức theo NĐ 30/2020/NĐ-CP. Add-in cho Word trên Windows và Mac. Không dùng AI, không gửi nội dung văn bản đi.</p>
<p><a class="btn" href="https://github.com/${REPO}/releases/latest">Tải add-in (zip, v${version})</a>
<a class="btn" href="manifest.xml">manifest.xml</a>
<a class="btn" href="vanban.html">Dùng bản web (không cần cài)</a></p>
<p>Hướng dẫn cài đặt cho Mac, Windows và cả cơ quan: <a href="https://github.com/${REPO}/blob/master/HUONG-DAN-CAI-DAT.md">HUONG-DAN-CAI-DAT.md</a>.</p>
<p><strong>Lưu ý:</strong> quy cách trình bày trong công cụ chưa được đối chiếu chính thức với văn bản luật. Hãy kiểm tra kết quả trước khi ban hành.</p>
</body></html>
`,
);

console.log(`docs/ đã sẵn sàng (host: ${baseUrl})`);
