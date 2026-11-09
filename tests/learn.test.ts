import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { beforeAll, describe, expect, it } from 'vitest';
import { readDocx, setXmlParser, setXmlSerializer, type DocFormat } from '../src/docx-read';
import type { BlockNode, DocumentIR, ParagraphNode } from '../src/ir';
import { checkAgainstProfile, learnProfile, planFixes, roleFor } from '../src/learn/profile';
import { renderDocx } from '../src/render-docx';

beforeAll(() => {
  setXmlParser((s) => new DOMParser().parseFromString(s, 'text/xml') as unknown as Document);
  setXmlSerializer((el) => new XMLSerializer().serializeToString(el as never));
});

/* ---- Dựng tài liệu thử bằng chính bộ dựng .docx của dự án (IR), độc lập với NĐ 30 ---- */

interface Fmt {
  align?: ParagraphNode['align'];
  size?: number;
  bold?: boolean;
  italic?: boolean;
  firstLineMm?: number;
  leftMm?: number;
  before?: number;
  line?: number;
}
const p = (text: string, f: Fmt = {}): ParagraphNode => {
  const run = { text, ...(f.bold ? { bold: true } : {}), ...(f.italic ? { italic: true } : {}) };
  const node: ParagraphNode = {
    type: 'paragraph',
    runs: [run],
    align: f.align ?? 'left',
    sizePt: f.size ?? 13,
    spaceBeforePt: f.before ?? 0,
    lineSpacing: f.line ?? 1,
  };
  if (f.firstLineMm) node.firstLineMm = f.firstLineMm;
  if (f.leftMm) node.leftMm = f.leftMm;
  return node;
};

/** Định dạng chuẩn của "giáo án" mẫu */
const STD = {
  title: { align: 'center', size: 16, bold: true } as Fmt,
  section: { size: 14, bold: true, before: 12 } as Fmt,
  sub: { size: 13, bold: true, italic: true, before: 6 } as Fmt,
  body: { align: 'justify', size: 13, firstLineMm: 10, line: 1.3, before: 3 } as Fmt,
  bullet: { size: 13, leftMm: 10, before: 2 } as Fmt,
  thead: { align: 'center', size: 12, bold: true } as Fmt,
  tcol0: { size: 12 } as Fmt,
  tcol1: { size: 12, italic: true } as Fmt,
};

function lessonPlan(f: typeof STD, margins = { top: 20, bottom: 20, left: 30, right: 20 }): DocumentIR {
  const cell = (blocks: BlockNode[], w: number) => ({ children: blocks, widthMm: w });
  const table: BlockNode = {
    type: 'table',
    rows: [
      [cell([p('Hoạt động của giáo viên', f.thead)], 80), cell([p('Hoạt động của học sinh', f.thead)], 80)],
      [cell([p('Giao nhiệm vụ cho các nhóm', f.tcol0)], 80), cell([p('Nhóm thảo luận và ghi kết quả', f.tcol1)], 80)],
      [cell([p('Quan sát, hỗ trợ học sinh', f.tcol0)], 80), cell([p('Đại diện nhóm trình bày', f.tcol1)], 80)],
    ],
  };
  return {
    fontFamily: 'Times New Roman',
    defaultSizePt: 13,
    page: {
      widthMm: 210,
      heightMm: 297,
      marginTopMm: margins.top,
      marginBottomMm: margins.bottom,
      marginLeftMm: margins.left,
      marginRightMm: margins.right,
      headerDistanceMm: 10,
      footerDistanceMm: 10,
      pageNumber: null,
    },
    children: [
      p('KẾ HOẠCH BÀI DẠY', f.title),
      p('I. MỤC TIÊU', f.section),
      p('1. Kiến thức:', f.sub),
      p('Học sinh nắm được khái niệm phân số và biết so sánh hai phân số cùng mẫu số trong các tình huống quen thuộc.', f.body),
      p('2. Năng lực:', f.sub),
      p('Phát triển năng lực tư duy, giải quyết vấn đề và hợp tác trong nhóm học tập.', f.body),
      p('II. CHUẨN BỊ', f.section),
      p('- Giáo viên: máy chiếu, phiếu học tập', f.bullet),
      p('- Học sinh: sách giáo khoa, vở ghi', f.bullet),
      p('III. TIẾN TRÌNH DẠY HỌC', f.section),
      table,
    ],
  };
}

const toDoc = async (ir: DocumentIR): Promise<DocFormat> => readDocx(await renderDocx(ir));
const byText = (doc: DocFormat, startsWith: string) => {
  const para = doc.paragraphs.find((x) => x.text.startsWith(startsWith));
  if (!para) throw new Error(`không có đoạn "${startsWith}"`);
  return para;
};

describe('đọc .docx', () => {
  it('đọc đúng lề, đoạn, định dạng và cột bảng', async () => {
    const doc = await toDoc(lessonPlan(STD));
    expect(doc.page?.marginLeftMm).toBeCloseTo(30, 0);
    expect(doc.page?.marginTopMm).toBeCloseTo(20, 0);
    const title = byText(doc, 'KẾ HOẠCH');
    expect(title.align).toBe('center');
    expect(title.runs[0]?.bold).toBe(true);
    expect(title.runs[0]?.sizePt).toBe(16);
    const body = byText(doc, 'Học sinh nắm');
    expect(body.align).toBe('justify');
    expect(body.firstLineMm).toBeCloseTo(10, 0);
    expect(body.lineSpacing).toBeCloseTo(1.3, 1);
    const cell = byText(doc, 'Nhóm thảo luận');
    expect(cell.table).toMatchObject({ row: 1, col: 1 });
    expect(doc.tables[0]?.colWidthsMm.map((x) => Math.round(x))).toEqual([80, 80]);
  });
});

describe('học định dạng từ file mẫu (không phải văn bản hành chính)', () => {
  it('gom các đoạn thành vai trò và học dấu hiệu nhận biết', async () => {
    const profile = learnProfile(await toDoc(lessonPlan(STD)), 'Giáo án mẫu');
    expect(profile.roles.length).toBeGreaterThanOrEqual(7);
    const title = profile.roles.find((r) => r.format.sizePt === 16);
    expect(title?.format).toMatchObject({ align: 'center', bold: true, caps: true });
    const section = profile.roles.find((r) => r.features.roman > 0.9);
    expect(section?.format).toMatchObject({ sizePt: 14, bold: true });
    const bullet = profile.roles.find((r) => r.features.bullet > 0.9);
    expect(bullet?.format.leftMm).toBeCloseTo(10, 0);
    const col1 = profile.roles.find((r) => r.tableCol === 1 && r.format.italic);
    expect(col1).toBeDefined();
    expect(profile.page?.marginLeftMm).toBeCloseTo(30, 0);
    expect(profile.font).toBe('Times New Roman');
  });

  it('chính file mẫu phải đạt chuẩn: không có lỗi nào', async () => {
    const doc = await toDoc(lessonPlan(STD));
    const profile = learnProfile(doc);
    const r = checkAgainstProfile(profile, doc);
    expect(r.issues, JSON.stringify(r.issues.slice(0, 2))).toEqual([]);
    expect(r.page).toEqual([]);
  });

  it('đoán đúng vai trò của các đoạn trong tài liệu khác (kể cả khi định dạng sai hết)', async () => {
    const profile = learnProfile(await toDoc(lessonPlan(STD)));
    const messy: typeof STD = Object.fromEntries(Object.keys(STD).map((k) => [k, { size: 11 }])) as unknown as typeof STD;
    const doc = await toDoc(lessonPlan(messy, { top: 25, bottom: 20, left: 30, right: 20 }));

    const roleName = (start: string) => roleFor(profile, byText(doc, start));
    expect(roleName('KẾ HOẠCH').format.sizePt).toBe(16);
    expect(roleName('III. TIẾN').format).toMatchObject({ sizePt: 14, bold: true });
    expect(roleName('1. Kiến thức').format).toMatchObject({ bold: true, italic: true });
    expect(roleName('- Giáo viên').format.leftMm).toBeCloseTo(10, 0);
    expect(roleName('Nhóm thảo luận').format.italic).toBe(true);
    expect(roleName('Giao nhiệm vụ').format.italic).toBe(false);
  });

  it('phát hiện lỗi định dạng và lề trang, đề xuất sửa đúng theo mẫu', async () => {
    const profile = learnProfile(await toDoc(lessonPlan(STD)));
    const messy: typeof STD = Object.fromEntries(Object.keys(STD).map((k) => [k, { size: 11 }])) as unknown as typeof STD;
    const doc = await toDoc(lessonPlan(messy, { top: 25, bottom: 20, left: 30, right: 20 }));

    const r = checkAgainstProfile(profile, doc);
    expect(r.issues.length).toBeGreaterThanOrEqual(8);
    expect(r.page.map((x) => x.field)).toEqual(['lề trên']);

    const fixes = planFixes(profile, doc);
    const heading = fixes.find((f) => f.text.startsWith('III. TIẾN'));
    expect(heading?.set).toMatchObject({ sizePt: 14, bold: true, spaceBeforePt: 12 });
    const body = fixes.find((f) => f.text.startsWith('Học sinh nắm'));
    expect(body?.set).toMatchObject({ align: 'justify', sizePt: 13 });
    expect(body?.set.lineSpacing).toBeCloseTo(1.3, 1);
    const title = fixes.find((f) => f.text.startsWith('KẾ HOẠCH'));
    expect(title?.set).toMatchObject({ align: 'center', sizePt: 16, bold: true });
  });
});
