import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { readDocx, setXmlParser, setXmlSerializer } from '../src/docx-read';
import { checkNd30, planNd30Fixes } from '../src/profiles/vn-nd30/check';
import { buildIR, assemble, classify, formatNd30, ND30_LAYOUT } from '../src/profiles/vn-nd30';
import type { Layout } from '../src/profiles/vn-nd30/build-ir';
import { renderDocx } from '../src/render-docx';

beforeAll(() => {
  setXmlParser((s) => new DOMParser().parseFromString(s, 'text/xml') as unknown as Document);
  setXmlSerializer((el) => new XMLSerializer().serializeToString(el as never));
});

const fixture = (n: string): string => readFileSync(new URL(`../fixtures/${n}.txt`, import.meta.url), 'utf8');

describe('kiểm tra thể thức NĐ 30', () => {
  for (const name of ['cong-van', 'quyet-dinh', 'bao-cao']) {
    it(`văn bản do công cụ tạo (${name}) tự đạt kiểm tra`, async () => {
      const { docx } = await formatNd30(fixture(name));
      const doc = await readDocx(docx);
      const r = checkNd30(doc);
      expect(r.page).toEqual([]);
      expect(r.tables).toEqual([]);
      expect(r.issues.map((i) => `${i.roleName}: ${i.text} -> ${i.diffs.map((d) => d.field).join(',')}`)).toEqual([]);
    });
  }
});

/** Tạo văn bản bằng một layout cố ý sai để thử bộ kiểm tra. */
async function wrongDoc(patch: (L: Layout) => void) {
  const L = JSON.parse(JSON.stringify(ND30_LAYOUT)) as Layout;
  patch(L);
  const bytes = await renderDocx(buildIR(assemble(classify(fixture('quyet-dinh'))), L));
  return readDocx(bytes);
}

describe('phát hiện lỗi định dạng', () => {
  it('lề trang và khổ giấy sai', async () => {
    const r = checkNd30(
      await wrongDoc((L) => {
        L.page.margins.leftMm = 15;
        L.page.margins.topMm = 10;
      }),
    );
    expect(r.page.map((p) => p.field).sort()).toEqual(['lề trái', 'lề trên']);
  });

  it('cỡ chữ nội dung nằm ngoài 13 đến 14 pt, và đề xuất sửa về cỡ gần nhất trong khoảng', async () => {
    const doc = await wrongDoc((L) => {
      L.defaults.bodySizePt = 16;
    });
    const r = checkNd30(doc);
    const bad = r.issues.filter((i) => i.diffs.some((d) => d.field === 'sizePt'));
    expect(bad.length).toBeGreaterThan(0);
    const plan = planNd30Fixes(r);
    const fixed = plan.items.filter((i) => i.set.sizePt !== undefined);
    expect(fixed.every((i) => i.set.sizePt === 14)).toBe(true);
  });

  it('căn lề nội dung không đều hai bên và không thụt đầu dòng', async () => {
    const r = checkNd30(
      await wrongDoc((L) => {
        L.body.paragraph.align = 'left';
        L.defaults.firstLineIndentMm = 0;
        L.defaults.firstLineIndentRangeMm = [0, 0];
      }),
    );
    const fields = new Set(r.issues.flatMap((i) => i.diffs.map((d) => d.field)));
    expect(fields.has('align')).toBe(true);
    expect(fields.has('firstLineMm')).toBe(true);
  });

  it('trong khoảng cho phép thì không báo lỗi (cỡ chữ 13 và thụt 10 mm đều hợp lệ)', async () => {
    const r = checkNd30(
      await wrongDoc((L) => {
        L.defaults.bodySizePt = 13;
        L.defaults.firstLineIndentMm = 10;
      }),
    );
    expect(r.issues).toEqual([]);
  });

  it('nhắc ngày tháng chưa thêm số 0', () => {
    const doc = {
      paragraphs: [
        { index: 0, text: 'Hà Nội, ngày 5 tháng 1 năm 2025', align: 'center', firstLineMm: 0, leftMm: 0, rightMm: 0, spaceBeforePt: 0, spaceAfterPt: 0, lineSpacing: 1, runs: [{ text: 'Hà Nội, ngày 5 tháng 1 năm 2025', bold: false, italic: true, caps: false, sizePt: 14 }] },
      ],
      tables: [],
    } as never;
    const r = checkNd30(doc);
    expect(r.notes.some((n) => n.includes('ngày 05 tháng 01 năm 2025'))).toBe(true);
  });
});
