/**
 * Kiểm tra một tài liệu Word đã soạn có đúng thể thức NĐ 30/2020/NĐ-CP không, và lập kế hoạch sửa định dạng tại chỗ.
 *
 * Cách làm: đọc định dạng (DocFormat) → phân loại từng đoạn thành nhãn thể thức (quốc hiệu, trích yếu, căn cứ, ...)
 * bằng bộ phân loại sẵn có → so sánh định dạng quan sát được với quy cách trong layout.json.
 * Quy cách NĐ 30 thường cho một khoảng (ví dụ cỡ chữ 13 đến 14), nên nằm trong khoảng là đạt.
 *
 * Chạy hoàn toàn trên máy, không dùng AI hay mạng. Chỉ sửa được định dạng; thiếu thành phần (số, nơi nhận ...)
 * thì báo ở phần ghi chú để người dùng bổ sung hoặc dùng chức năng "Chuẩn hoá" tạo văn bản mới.
 */
import type { DocFormat, ParaFmt, TableBorders } from '../../docx-read';
import {
  bordersEqual,
  describeBorders,
  observeFormat,
  type Align,
  type CheckResult,
  type FieldDiff,
  type FixItem,
  type PageIssue,
  type ParagraphIssue,
  type RoleFormat,
  type TableFix,
  type TableIssue,
} from '../../learn/profile';
import { assemble } from './assemble';
import { ND30_LAYOUT, type Layout, type Style } from './build-ir';
import { classifyParagraphs } from './classify';
import { LABEL_NAMES, type Label } from './types';

/** Quy cách mong đợi của một nhãn. Trường vắng mặt = không kiểm tra. */
interface Expect {
  align?: Align;
  size?: [number, number];
  bold?: boolean;
  italic?: boolean;
  caps?: boolean;
  firstLine?: [number, number];
  lineSpacing?: [number, number];
  /** Khoảng cách tối thiểu giữa các đoạn (trước + sau), pt. */
  gapMinPt?: number;
}

const fromStyle = (st: Style): Expect => {
  const e: Expect = {};
  if (st.align) e.align = st.align;
  const size = st.sizeRangePt ?? (st.sizePt !== undefined ? [st.sizePt, st.sizePt] : undefined);
  if (size) e.size = [size[0], size[1]];
  if (st.bold !== undefined) e.bold = st.bold;
  if (st.italic !== undefined) e.italic = st.italic;
  if (st.caps) e.caps = true;
  return e;
};

/** Nhãn nằm trong khối đầu hoặc khối cuối (dàn bằng bảng không viền). */
const LAYOUT_TABLE_LABELS = new Set<Label>([
  'parent-org',
  'issuer',
  'national-title',
  'motto',
  'number',
  'subject-cv',
  'place-date',
  'recipients-title',
  'recipient',
  'signer-title',
  'signer-name',
]);

export function expectedFor(label: Label, ctx: { docType?: string }, L: Layout): Expect | null {
  const D = L.defaults;
  const body: Expect = {
    align: L.body.paragraph.align ?? 'justify',
    size: D.bodySizeRangePt ?? [D.bodySizePt, D.bodySizePt],
    firstLine: D.firstLineIndentRangeMm ?? [D.firstLineIndentMm, D.firstLineIndentMm],
    lineSpacing: D.lineSpacingRange ?? [D.lineSpacing, D.lineSpacing],
    gapMinPt: D.paragraphSpaceBeforePt,
  };
  const heading = (level: string): Expect => {
    const st = L.body.headings[level] ?? {};
    const e: Expect = { size: body.size as [number, number] };
    if (st.bold !== undefined) e.bold = st.bold;
    if (st.italic !== undefined) e.italic = st.italic;
    return e;
  };
  switch (label) {
    case 'parent-org':
      return fromStyle(L.header.parentOrg);
    case 'issuer':
      return fromStyle(L.header.issuer);
    case 'national-title':
      return fromStyle(L.header.nationalTitle);
    case 'motto':
      return fromStyle(L.header.motto);
    case 'number':
      return fromStyle(L.header.number);
    case 'subject-cv':
      return fromStyle(L.header.subjectCongVan);
    case 'place-date':
      return fromStyle(L.header.placeDate);
    case 'doc-type':
      return fromStyle(L.title.docTypeName);
    case 'subject':
      return fromStyle(L.title.subject);
    case 'addressee':
      return fromStyle(L.addressee);
    case 'authority':
      return fromStyle(L.body.authority);
    case 'enact':
      return fromStyle(L.body.enact);
    case 'attachment-note':
      return fromStyle(L.body.note);
    case 'chapter':
    case 'chapter-title':
      return fromStyle(L.body.chapter);
    case 'basis': {
      const e: Expect = { ...body };
      if (ctx.docType && L.body.basis.italicFor.includes(ctx.docType)) e.italic = true;
      return e;
    }
    case 'paragraph':
      return body;
    case 'article':
      return { size: body.size as [number, number] };
    case 'heading1':
      return heading('1');
    case 'heading2':
      return heading('2');
    case 'heading3':
      return heading('3');
    case 'heading4':
      return heading('4');
    case 'recipients-title':
      return fromStyle(L.closing.recipientsTitle);
    case 'recipient':
      return fromStyle(L.closing.recipientItem);
    case 'signer-title':
      return { ...fromStyle(L.closing.signer), caps: true };
    case 'signer-name':
      return fromStyle(L.closing.signer);
    default:
      return null; // decor, signature-note: không kiểm tra
  }
}

const TOL = { size: 0.25, mm: 1.5, line: 0.08, pt: 1.5 };

/** Giá trị gần nhất nằm trong khoảng [lo, hi]. */
const clamp = (x: number, [lo, hi]: [number, number]): number => Math.min(hi, Math.max(lo, x));
const inRange = (x: number, [lo, hi]: [number, number], tol: number): boolean => x >= lo - tol && x <= hi + tol;
const rangeText = ([lo, hi]: [number, number], unit: string): string => (lo === hi ? `${lo} ${unit}` : `${lo} đến ${hi} ${unit}`);

export function diffExpect(e: Expect, a: RoleFormat, font: string, inTable: boolean): FieldDiff[] {
  const out: FieldDiff[] = [];
  const add = (field: FieldDiff['field'], expected: FieldDiff['expected'], actual: FieldDiff['actual'], expectedLabel?: string): void => {
    out.push(expectedLabel ? { field, expected, actual, expectedLabel } : { field, expected, actual });
  };
  if (e.align && a.align !== e.align) add('align', e.align, a.align);
  if (e.size && !inRange(a.sizePt, e.size, TOL.size)) add('sizePt', clamp(a.sizePt, e.size), a.sizePt, rangeText(e.size, 'pt'));
  if (e.bold !== undefined && a.bold !== e.bold) add('bold', e.bold, a.bold);
  if (e.italic !== undefined && a.italic !== e.italic) add('italic', e.italic, a.italic);
  if (e.caps && !a.caps) add('caps', true, false);
  if (!inTable) {
    if (e.firstLine && !inRange(a.firstLineMm, e.firstLine, TOL.mm))
      add('firstLineMm', clamp(a.firstLineMm, e.firstLine), a.firstLineMm, rangeText(e.firstLine, 'mm'));
    if (e.lineSpacing && !inRange(a.lineSpacing, e.lineSpacing, TOL.line))
      add('lineSpacing', clamp(a.lineSpacing, e.lineSpacing), a.lineSpacing, `×${e.lineSpacing[0]} đến ×${e.lineSpacing[1]}`);
    if (e.gapMinPt !== undefined && a.spaceBeforePt + a.spaceAfterPt < e.gapMinPt - TOL.pt)
      add('spaceBeforePt', e.gapMinPt - a.spaceAfterPt, a.spaceBeforePt, `ít nhất ${e.gapMinPt} pt giữa hai đoạn`);
  }
  if (a.font && a.font.toLowerCase() !== font.toLowerCase()) add('font', font, a.font);
  return out;
}

export interface Nd30Check extends CheckResult {
  /** Nhãn thể thức đã đoán cho từng đoạn không trống (theo chỉ số đoạn). */
  labelOf: Map<number, Label>;
  /** Việc cần người dùng tự làm: thiếu thành phần, sai cách viết ngày ... (không sửa tự động được). */
  notes: string[];
  /** Sửa khung bảng của các bảng dàn trang (phần đầu, nơi nhận / chữ ký): phải không viền. */
  tableFixes: TableFix[];
}

const NO_BORDERS: TableBorders = { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };

const RE_DATE = /ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})/iu;

/** Ngày dưới 10 và tháng 1, 2 phải thêm số 0 ở trước (Điều 8 NĐ 30 / Phụ lục I). */
function dateNote(text: string): string | null {
  const m = RE_DATE.exec(text);
  if (!m) return null;
  const [d, mo] = [m[1] as string, m[2] as string];
  const wantD = Number(d) < 10 ? `0${Number(d)}` : d;
  const wantM = Number(mo) <= 2 ? `0${Number(mo)}` : mo;
  if (d === wantD && mo === wantM) return null;
  return `Ngày tháng cần viết “ngày ${wantD} tháng ${wantM} năm ${m[3]}” (số ngày dưới 10 và tháng 1, 2 thêm số 0 ở trước). Đang viết “${m[0]}”. Hãy tự sửa chữ này.`;
}

export function checkNd30(doc: DocFormat, L: Layout = ND30_LAYOUT): Nd30Check {
  const labeled = classifyParagraphs(doc.paragraphs.map((p) => p.text));
  const model = assemble(labeled);
  const labelOf = new Map<number, Label>();
  const issues: ParagraphIssue[] = [];
  const notes: string[] = [];
  const byTable = new Map<number, Label[]>();

  for (const l of labeled) {
    const p = doc.paragraphs[l.para] as ParaFmt;
    labelOf.set(p.index, l.label);
    if (p.table) byTable.set(p.table.id, [...(byTable.get(p.table.id) ?? []), l.label]);
    const e = expectedFor(l.label, model.docType ? { docType: model.docType } : {}, L);
    if (!e) continue;
    const diffs = diffExpect(e, observeFormat(p, doc), L.font, !!p.table);
    if (diffs.length)
      issues.push({ index: p.index, text: p.text.trim().slice(0, 70), roleId: l.label, roleName: LABEL_NAMES[l.label], diffs });
    if (l.label === 'place-date') {
      const n = dateNote(p.text);
      if (n) notes.push(n);
    }
  }

  // Trang
  const page: PageIssue[] = [];
  if (doc.page) {
    const P = L.page;
    const cmp = (name: string, want: number, got: number, tol: number): void => {
      if (Math.abs(want - got) > tol) page.push({ field: name, expected: `${want} mm`, actual: `${Math.round(got * 10) / 10} mm` });
    };
    cmp('khổ giấy (rộng)', P.widthMm, doc.page.widthMm, 1.5);
    cmp('khổ giấy (cao)', P.heightMm, doc.page.heightMm, 1.5);
    const r = P.margins.range;
    const margin = (name: string, want: number, got: number, range?: [number, number]): void => {
      if (range ? got < range[0] - 0.7 || got > range[1] + 0.7 : Math.abs(want - got) > 0.7)
        page.push({
          field: name,
          expected: range ? `${range[0]} đến ${range[1]} mm` : `${want} mm`,
          actual: `${Math.round(got * 10) / 10} mm`,
        });
    };
    margin('lề trên', P.margins.topMm, doc.page.marginTopMm, r?.topMm);
    margin('lề dưới', P.margins.bottomMm, doc.page.marginBottomMm, r?.bottomMm);
    margin('lề trái', P.margins.leftMm, doc.page.marginLeftMm, r?.leftMm);
    margin('lề phải', P.margins.rightMm, doc.page.marginRightMm, r?.rightMm);
  }
  if (doc.hasPageNumber === false && P_NEEDS_NUMBER(L))
    notes.push('Chưa thấy đánh số trang. Nếu văn bản dài hơn một trang, cần đánh số ở giữa lề trên và bỏ số ở trang đầu (Chèn > Số trang).');

  // Bảng dàn trang: không viền
  const tableIssues: TableIssue[] = [];
  const tableFixes: TableFix[] = [];
  doc.tables.forEach((t, tableIndex) => {
    const labels = byTable.get(t.id);
    if (!labels?.length || !labels.every((x) => LAYOUT_TABLE_LABELS.has(x))) return;
    if (bordersEqual(NO_BORDERS, t.borders)) return;
    const cols = t.colWidthsMm.length;
    tableIssues.push({
      tableIndex,
      cols,
      diffs: [{ field: 'borders', expected: 'không viền', actual: describeBorders(t.borders) }],
    });
    tableFixes.push({ tableIndex, cols, set: { borders: NO_BORDERS } });
  });

  // Thành phần còn thiếu (lấy từ bộ dựng mô hình)
  for (const w of model.warnings) notes.push(w);

  return {
    paragraphs: labelOf.size,
    issues,
    page,
    tables: tableIssues,
    unmatchedTables: [],
    roleOf: new Map([...labelOf].map(([i, l]) => [i, l])),
    labelOf,
    notes,
    tableFixes,
  };
}

const P_NEEDS_NUMBER = (L: Layout): boolean => !!L.page.pageNumber;

/** Kế hoạch sửa định dạng đoạn: đặt lại các thuộc tính lệch về giá trị chuẩn. */
export function planNd30Fixes(check: Nd30Check): { items: FixItem[]; tableItems: TableFix[] } {
  const items = check.issues.map((i) => {
    const set: Partial<RoleFormat> = {};
    for (const d of i.diffs) (set as Record<string, unknown>)[d.field] = d.expected;
    return { index: i.index, text: i.text, roleId: i.roleId, set };
  });
  return { items, tableItems: check.tableFixes };
}
