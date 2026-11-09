// Sinh biểu tượng PNG (16/32/80 px) không cần thư viện ngoài. Chạy: node scripts/make-icons.mjs
import { mkdir, writeFile } from 'node:fs/promises';
import { crc32, deflateSync } from 'node:zlib';

function png(size, pixel) {
  const raw = Buffer.alloc((size * 4 + 1) * size);
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0; // filter: none
    for (let x = 0; x < size; x++) {
      const [r, g, b, a] = pixel(x + 0.5, y + 0.5, size);
      raw.set([r, g, b, a], y * (size * 4 + 1) + 1 + x * 4);
    }
  }
  const chunk = (type, data) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(td) >>> 0);
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr.set([8, 6, 0, 0, 0], 8); // 8-bit RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

const BLUE = [26, 86, 196, 255];
const WHITE = [255, 255, 255, 255];
const CLEAR = [0, 0, 0, 0];

// Nền xanh bo góc, tờ giấy trắng ở giữa, ba dòng chữ xanh.
function icon(x, y, s) {
  const u = x / s;
  const v = y / s;
  const r = 0.18;
  const cx = Math.min(Math.max(u, r), 1 - r);
  const cy = Math.min(Math.max(v, r), 1 - r);
  if ((u - cx) ** 2 + (v - cy) ** 2 > r * r) return CLEAR;
  const inPage = u > 0.26 && u < 0.74 && v > 0.16 && v < 0.84;
  if (!inPage) return BLUE;
  const lines = [0.32, 0.46, 0.6, 0.74];
  for (const [i, ly] of lines.entries()) {
    const right = i === 0 ? 0.62 : i === 3 ? 0.55 : 0.66;
    if (Math.abs(v - ly) < 0.035 && u > 0.34 && u < right) return BLUE;
  }
  return WHITE;
}

await mkdir('assets', { recursive: true });
for (const s of [16, 32, 80]) {
  await writeFile(`assets/icon-${s}.png`, png(s, icon));
  console.log(`assets/icon-${s}.png`);
}
