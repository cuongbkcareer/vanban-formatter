// Cấu hình dùng chung cho các script đóng gói.
import { readFile } from 'node:fs/promises';

export const ADDIN_ID = '81a47bdf-3c91-4d65-9322-112a4078e6c5'; // GUID cố định: KHÔNG đổi giữa các phiên bản
export const REPO = 'cuongbkcareer/vanban-formatter';
export const DEFAULT_BASE_URL = 'https://cuongbkcareer.github.io/vanban-formatter';
export const SUPPORT_URL = `https://github.com/${REPO}/issues`;

export async function readVersion() {
  const pkg = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));
  return pkg.version;
}
