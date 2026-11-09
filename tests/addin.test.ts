import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { generateManifest } from '../scripts/manifest.mjs';
import { cleanWordText, wordHost } from '../web/word-host';

const BASE = 'https://example.github.io/vanban-formatter';

describe('manifest.xml', () => {
  const xml = generateManifest({ baseUrl: BASE + '/', version: '1.2.3' });

  it('đúng loại add-in cho Word, có id cố định và phiên bản 4 phần', () => {
    expect(xml).toContain('xsi:type="TaskPaneApp"');
    expect(xml).toContain('<Host Name="Document"/>');
    expect(xml).toMatch(/<Id>[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}<\/Id>/);
    expect(xml).toContain('<Version>1.2.3.0</Version>');
    expect(xml).toContain('<Permissions>ReadWriteDocument</Permissions>');
    expect(xml).toContain('<Set Name="WordApi" MinVersion="1.3"/>');
  });

  it('mọi địa chỉ đều là HTTPS thuộc host đã chọn', () => {
    const urls = [...xml.matchAll(/DefaultValue="(https?:[^"]+)"/g)].map((m) => m[1]!);
    expect(urls.length).toBeGreaterThanOrEqual(6);
    for (const u of urls) {
      expect(u.startsWith('https://')).toBe(true);
    }
    expect(xml).toContain(`<SourceLocation DefaultValue="${BASE}/taskpane.html"/>`);
    expect(xml).toContain('<AppDomain>https://example.github.io</AppDomain>');
  });

  it('mọi resid đều có định nghĩa tương ứng trong Resources', () => {
    const used = new Set([...xml.matchAll(/resid="([^"]+)"/g)].map((m) => m[1]!));
    const defined = new Set([...xml.matchAll(/<bt:(?:Image|Url|String) id="([^"]+)"/g)].map((m) => m[1]!));
    for (const r of used) expect(defined.has(r), `thiếu resource ${r}`).toBe(true);
  });

  it('có đủ các phần của mẫu manifest chuẩn: GetStarted, FunctionFile, nhóm trên thẻ Home', () => {
    expect(xml).toContain('<FunctionFile resid="Commands.Url"/>');
    expect(xml).toContain('<GetStarted>');
    expect(xml).toContain('<OfficeTab id="TabHome">');
    expect(xml).toContain('<Action xsi:type="ShowTaskpane">');
  });

  it('mọi resid và id tài nguyên đều không quá 32 ký tự (giới hạn của Office)', () => {
    const ids = [...xml.matchAll(/(?:resid|<bt:(?:Image|Url|String) id)="([^"]+)"/g)].map((m) => m[1]!);
    expect(ids.length).toBeGreaterThan(10);
    for (const id of ids) expect(id.length, id).toBeLessThanOrEqual(32);
  });

  it('từ chối host không dùng HTTPS', () => {
    expect(() => generateManifest({ baseUrl: 'http://x.test', version: '1.0.0' })).toThrow(/HTTPS/);
  });
});

describe('cầu nối Word (Office.js giả lập)', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  function stubWord(opts: { selectionText: string; bodyTexts: string[]; selTexts?: string[]; sets?: string[] }) {
    const calls: { createDocument?: string; opened?: boolean; insert?: unknown[] } = {};
    const paragraphs = (texts: string[]) => ({ items: texts.map((text) => ({ text })), load: vi.fn() });
    const ctx = {
      document: {
        getSelection: () => ({ text: opts.selectionText, load: vi.fn(), paragraphs: paragraphs(opts.selTexts ?? []) }),
        body: {
          paragraphs: paragraphs(opts.bodyTexts),
          insertFileFromBase64: (...args: unknown[]) => {
            calls.insert = args;
          },
        },
      },
      application: {
        createDocument: (b64: string) => {
          calls.createDocument = b64;
          return {
            open: () => {
              calls.opened = true;
            },
          };
        },
      },
      sync: vi.fn(async () => {}),
    };
    vi.stubGlobal('Word', { run: async (fn: (c: typeof ctx) => Promise<unknown>) => fn(ctx) });
    vi.stubGlobal('Office', {
      context: { requirements: { isSetSupported: (n: string, v: string) => (opts.sets ?? []).includes(`${n} ${v}`) } },
    });
    return calls;
  }

  it('đọc vùng chọn nếu có, ngược lại đọc toàn bộ tài liệu', async () => {
    stubWord({ selectionText: 'abc', selTexts: ['dòng chọn 1', 'dòng chọn 2'], bodyTexts: ['toàn bộ'] });
    expect(await wordHost.readSource!()).toBe('dòng chọn 1\ndòng chọn 2');
    stubWord({ selectionText: '  ', bodyTexts: ['một', 'hai'] });
    expect(await wordHost.readSource!()).toBe('một\nhai');
  });

  it('làm sạch ký tự điều khiển của Word', () => {
    expect(cleanWordText('Số: 1\u0007')).toBe('Số: 1');
    expect(cleanWordText('a\u000bb\rc')).toBe('a\nb\nc');
  });

  it('mở tài liệu mới từ chuỗi base64', async () => {
    const calls = stubWord({ selectionText: '', bodyTexts: [] });
    await wordHost.openNew!('QUJD');
    expect(calls.createDocument).toBe('QUJD');
    expect(calls.opened).toBe(true);
  });

  it('thay nội dung: có tuỳ chọn khi WordApi 1.5, không có khi chỉ 1.3', async () => {
    const c15 = stubWord({ selectionText: '', bodyTexts: [], sets: ['WordApi 1.5'] });
    await wordHost.replaceCurrent!('QUJD');
    expect(c15.insert?.[1]).toBe('Replace');
    expect(c15.insert?.[2]).toMatchObject({ importStyles: false, importParagraphSpacing: true });
    const c13 = stubWord({ selectionText: '', bodyTexts: [] });
    await wordHost.replaceCurrent!('QUJD');
    expect(c13.insert).toHaveLength(2);
  });
});

describe('cam kết bảo mật: mã của dự án không gọi mạng', () => {
  const FORBIDDEN = /\b(fetch\s*\(|XMLHttpRequest|WebSocket|EventSource|sendBeacon|navigator\.connection|importScripts)/;
  function files(dir: string): string[] {
    return readdirSync(dir).flatMap((f) => {
      const p = join(dir, f);
      return statSync(p).isDirectory() ? files(p) : /\.(ts|mjs|html)$/.test(p) ? [p] : [];
    });
  }
  it('src/ và web/ không dùng API mạng', () => {
    const offenders = [...files('src'), ...files('web')].filter((f) => FORBIDDEN.test(readFileSync(f, 'utf8')));
    expect(offenders).toEqual([]);
  });
});
