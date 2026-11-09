import { spawnSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { generateManifest } from '../scripts/manifest.mjs';

// Chạy script thật nhưng chỉ chép vào thư mục tạm (VANBAN_WEF_DIR), không đụng tới Word hay Finder.
const SCRIPT = new URL('../installer/cai-dat-mac.sh', import.meta.url).pathname;
let dir: string;

function run(args: string[], env: Record<string, string> = {}) {
  return spawnSync('bash', [join(dir, 'cai-dat-mac.sh'), ...args], {
    encoding: 'utf8',
    env: { ...process.env, VANBAN_WEF_DIR: join(dir, 'wef'), ...env },
  });
}

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'vanban-install-'));
  copyFileSync(SCRIPT, join(dir, 'cai-dat-mac.sh'));
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

describe('cai-dat-mac.sh', () => {
  it('cú pháp bash hợp lệ', () => {
    expect(spawnSync('bash', ['-n', SCRIPT]).status).toBe(0);
  });

  it('cài từ manifest.xml cạnh script rồi gỡ được', () => {
    const manifest = generateManifest({ baseUrl: 'https://example.test/x', version: '1.0.0' });
    spawnSync('bash', ['-c', `cat > "${join(dir, 'manifest.xml')}"`], { input: manifest });
    const r = run([]);
    expect(r.status, r.stderr).toBe(0);
    expect(r.stdout).toContain('Đã cài add-in');
    const installed = join(dir, 'wef', 'vanban-formatter.manifest.xml');
    expect(readFileSync(installed, 'utf8')).toBe(manifest);

    const u = run(['--uninstall']);
    expect(u.status).toBe(0);
    expect(existsSync(installed)).toBe(false);
  });

  it('từ chối manifest sai (thiếu mã add-in)', () => {
    spawnSync('bash', ['-c', `echo '<xml/>' > "${join(dir, 'manifest.xml')}"`]);
    const r = run([]);
    expect(r.status).not.toBe(0);
    expect(r.stderr).toContain('không đúng');
    expect(existsSync(join(dir, 'wef', 'vanban-formatter.manifest.xml'))).toBe(false);
  });

  it('không có manifest và không tải được thì báo lỗi rõ ràng', () => {
    const r = run([], { VANBAN_BASE_URL: 'https://127.0.0.1:9' });
    expect(r.status).not.toBe(0);
    expect(r.stderr).toContain('Không tải được manifest.xml');
  });
});
