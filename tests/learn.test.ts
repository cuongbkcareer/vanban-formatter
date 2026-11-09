import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { beforeAll, describe, expect, it } from 'vitest';
import { parseDocumentXml, readDocx, setXmlParser, setXmlSerializer, type DocFormat } from '../src/docx-read';
import type { BlockNode, DocumentIR, ParagraphNode, TableFormatSpec } from '../src/ir';
import { checkAgainstProfile, learnProfile, planFixes, planTableFixes, roleFor } from '../src/learn/profile';
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

const B = { sizePt: 0.5, color: '000000' };
/** Khung bảng chuẩn: viền đơn mọi cạnh, hàng tiêu đề tô xám và lặp lại, căn giữa theo chiều dọc */
const TABLE_STD: TableFormatSpec = {
  borders: { top: B, bottom: B, left: B, right: B, insideH: B, insideV: B },
  headerShading: 'D9D9D9',
  headerRepeat: true,
  vAlign: 'center',
  cellMarginMm: [1, 2, 1, 2],
};

function lessonPlan(
  f: typeof STD,
  margins = { top: 20, bottom: 20, left: 30, right: 20 },
  tbl: { format?: TableFormatSpec; widths?: [number, number] } = { format: TABLE_STD, widths: [60, 100] },
): DocumentIR {
  const [w0, w1] = tbl.widths ?? [80, 80];
  const cell = (blocks: BlockNode[], w: number) => ({ children: blocks, widthMm: w });
  const table: BlockNode = {
    type: 'table',
    ...(tbl.format ? { format: tbl.format } : {}),
    rows: [
      [cell([p('Hoạt động của giáo viên', f.thead)], w0), cell([p('Hoạt động của học sinh', f.thead)], w1)],
      [cell([p('Giao nhiệm vụ cho các nhóm', f.tcol0)], w0), cell([p('Nhóm thảo luận và ghi kết quả', f.tcol1)], w1)],
      [cell([p('Quan sát, hỗ trợ học sinh', f.tcol0)], w0), cell([p('Đại diện nhóm trình bày', f.tcol1)], w1)],
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
    expect(doc.tables[0]?.colWidthsMm.map((x) => Math.round(x))).toEqual([60, 100]);
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

describe('khung bảng giáo án', () => {
  it('đọc viền, tô nền hàng tiêu đề, lặp hàng tiêu đề, căn dọc, lề ô', async () => {
    const t = (await toDoc(lessonPlan(STD))).tables[0]!;
    expect(t.rows).toBe(3);
    expect(t.borders.top).toEqual({ sizePt: 0.5, color: '000000', style: 'single' });
    expect(t.borders.insideV).toEqual({ sizePt: 0.5, color: '000000', style: 'single' });
    expect(t.headerShading).toBe('D9D9D9');
    expect(t.headerRepeat).toBe(true);
    expect(t.vAlign).toBe('center');
    expect(t.cellMarginMm?.map((x) => Math.round(x))).toEqual([1, 2, 1, 2]);
  });

  it('bảng không viền, không tô nền được đọc là "không"', async () => {
    const t = (await toDoc(lessonPlan(STD, undefined, { widths: [80, 80] }))).tables[0]!;
    expect(t.borders).toEqual({ top: null, bottom: null, left: null, right: null, insideH: null, insideV: null });
    expect(t.headerShading).toBeNull();
    expect(t.headerRepeat).toBe(false);
  });

  it('học khung bảng theo số cột, độ rộng lưu theo tỉ lệ', async () => {
    const profile = learnProfile(await toDoc(lessonPlan(STD)));
    expect(profile.tables).toHaveLength(1);
    const spec = profile.tables![0]!;
    expect(spec.cols).toBe(2);
    expect(spec.colRatios).toEqual([0.38, 0.62]);
    expect(spec.headerShading).toBe('D9D9D9');
    expect(spec.borders.bottom).not.toBeNull();
  });

  it('file mẫu đạt chuẩn: không có lỗi bảng', async () => {
    const doc = await toDoc(lessonPlan(STD));
    const r = checkAgainstProfile(learnProfile(doc), doc);
    expect(r.tables).toEqual([]);
    expect(r.unmatchedTables).toEqual([]);
  });

  it('bảng sai khung (không viền, cột 50/50, không tô nền) bị báo đủ và có kế hoạch sửa', async () => {
    const profile = learnProfile(await toDoc(lessonPlan(STD)));
    const bad = await toDoc(lessonPlan(STD, undefined, { widths: [80, 80] }));
    const r = checkAgainstProfile(profile, bad);
    expect(r.tables).toHaveLength(1);
    const fields = r.tables[0]!.diffs.map((d) => d.field).sort();
    expect(fields).toEqual(['borders', 'cellMarginMm', 'colRatios', 'headerRepeat', 'headerShading', 'vAlign']);
    expect(r.tables[0]!.diffs.find((d) => d.field === 'borders')).toMatchObject({ expected: 'viền mọi cạnh 0.5 pt #000000', actual: 'không viền' });

    const fixes = planTableFixes(profile, bad);
    expect(fixes).toHaveLength(1);
    expect(fixes[0]!.set).toMatchObject({ headerShading: 'D9D9D9', headerRepeat: true, vAlign: 'center', colRatios: [0.38, 0.62] });
    expect(fixes[0]!.set.borders?.top).not.toBeNull();
  });

  it('bảng có số cột khác mẫu thì không đoán bừa: chỉ báo "chưa có mẫu"', async () => {
    const profile = learnProfile(await toDoc(lessonPlan(STD)));
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const cell = '<w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>';
    const doc = parseDocumentXml(
      `<w:document ${NS}><w:body><w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell}${cell}${cell}</w:tr></w:tbl></w:body></w:document>`,
    );
    const r = checkAgainstProfile(profile, doc);
    expect(r.tables).toEqual([]);
    expect(r.unmatchedTables).toEqual([0]);
  });

  it('đọc viền từ Style bảng có kế thừa (ví dụ "Table Grid" của Word)', () => {
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const styles = `<w:styles ${NS}>
      <w:style w:type="table" w:styleId="TableNormal"><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
      <w:style w:type="table" w:styleId="TableGrid"><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders>
        <w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/>
        <w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/>
        <w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>`;
    const cell = '<w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>';
    const doc = parseDocumentXml(
      `<w:document ${NS}><w:body><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell}${cell}</w:tr></w:tbl></w:body></w:document>`,
      { stylesXml: styles },
    );
    const t = doc.tables[0]!;
    expect(t.borders.top).toEqual({ sizePt: 0.5, color: '000000', style: 'single' });
    expect(t.borders.insideV).not.toBeNull();
    expect(t.cellMarginMm?.[3]).toBeCloseTo(1.9, 1);
  });
});
