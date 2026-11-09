/**
 * Học định dạng từ một file mẫu chuẩn (bất kỳ loại văn bản nào: văn bản hành chính, giáo án, báo cáo, ...).
 *
 *  1. Gom các đoạn của file mẫu thành "vai trò" theo kiểu định dạng giống nhau
 *     (căn lề, cỡ chữ, đậm/nghiêng, thụt lề, giãn dòng, nằm trong bảng ở cột nào ...).
 *  2. Với mỗi vai trò, học dấu hiệu của nội dung: có bắt đầu bằng "I." / "1." / "a)" / gạch đầu dòng không,
 *     có viết hoa toàn bộ không, có kết thúc bằng dấu hai chấm không, dài hay ngắn.
 *  3. Với tài liệu khác, đoán vai trò của từng đoạn (Bayes đơn giản trên các dấu hiệu đó),
 *     rồi so sánh hoặc đề xuất sửa định dạng theo đúng vai trò.
 *
 * Không dùng AI hay mạng. Cấu trúc nội dung hoàn toàn mới vẫn phụ thuộc vào việc dấu hiệu văn bản có phân biệt
 * được các vai trò hay không; kết quả luôn cho người dùng xem lại.
 */
import type { DocFormat, PageFmt, ParaFmt } from '../docx-read';
import { isAllCaps, wordCount } from '../text';

export type Align = 'left' | 'center' | 'right' | 'justify';

export interface RoleFormat {
  align: Align;
  sizePt: number;
  bold: boolean;
  italic: boolean;
  caps: boolean;
  font?: string;
  firstLineMm: number;
  leftMm: number;
  spaceBeforePt: number;
  spaceAfterPt: number;
  lineSpacing: number;
}

/** Tỉ lệ (0..1) số đoạn của vai trò có từng dấu hiệu. */
export interface RoleFeatures {
  roman: number;
  decimal: number;
  letter: number;
  bullet: number;
  caps: number;
  colon: number;
  listed: number;
  avgWords: number;
}

export interface Role {
  id: string;
  /** Tên gọi: tự đặt, người dùng có thể đổi. */
  name: string;
  count: number;
  format: RoleFormat;
  /** Tỉ lệ đoạn của vai trò nằm trong bảng, và cột phổ biến nhất (nếu có). */
  inTable: number;
  /** Tỉ lệ đoạn nằm ở hàng đầu của bảng (hàng tiêu đề). */
  firstRow: number;
  tableCol?: number;
  features: RoleFeatures;
  examples: string[];
}

export interface PageSpec {
  widthMm: number;
  heightMm: number;
  marginTopMm: number;
  marginBottomMm: number;
  marginLeftMm: number;
  marginRightMm: number;
}

export interface FormatProfile {
  version: 1;
  name: string;
  createdFrom?: string;
  page?: PageSpec;
  font?: string;
  roles: Role[];
}

/* ------------------------------------------------------------------ */
/* Đặc trưng của một đoạn                                              */
/* ------------------------------------------------------------------ */

const RE_ROMAN = /^[IVXLC]+\s*[.)]\s*\S/;
const RE_DECIMAL = /^\d+(?:[.)]|\.\d+)/;
const RE_LETTER = /^[a-zđA-ZĐ]\s*\)/;
const RE_BULLET = /^[-–—+•*●○▪■□✓➢➤]\s*\S/u;

export interface ParaFeatures {
  roman: boolean;
  decimal: boolean;
  letter: boolean;
  bullet: boolean;
  caps: boolean;
  colon: boolean;
  listed: boolean;
  words: number;
}

export function featuresOf(p: ParaFmt): ParaFeatures {
  const t = p.text.trim();
  return {
    roman: RE_ROMAN.test(t),
    decimal: RE_DECIMAL.test(t),
    letter: RE_LETTER.test(t),
    bullet: RE_BULLET.test(t),
    caps: isAllCaps(t) && wordCount(t) >= 1 && /\p{L}{2,}/u.test(t),
    colon: /:\s*$/.test(t),
    listed: p.listLevel !== undefined,
    words: wordCount(t),
  };
}

const round = (x: number, step: number): number => Math.round(x / step) * step;

/** Chỉ số cỡ chữ chiếm nhiều ký tự nhất; không có thì dùng cỡ mặc định của tài liệu. */
function dominantSize(p: ParaFmt, fallback: number): number {
  const w = new Map<number, number>();
  for (const r of p.runs) {
    const s = r.sizePt ?? fallback;
    w.set(s, (w.get(s) ?? 0) + r.text.length);
  }
  let best = fallback;
  let max = -1;
  for (const [s, n] of w) if (n > max) [best, max] = [s, n];
  return best;
}

function majority(p: ParaFmt, pick: (r: ParaFmt['runs'][number]) => boolean): boolean {
  let yes = 0;
  let all = 0;
  for (const r of p.runs) {
    const n = r.text.replace(/\s/g, '').length;
    all += n;
    if (pick(r)) yes += n;
  }
  return all > 0 && yes / all > 0.5;
}

function dominantFont(p: ParaFmt, fallback?: string): string | undefined {
  const w = new Map<string, number>();
  for (const r of p.runs) {
    const f = r.font ?? fallback;
    if (f) w.set(f, (w.get(f) ?? 0) + r.text.length);
  }
  let best: string | undefined;
  let max = -1;
  for (const [f, n] of w) if (n > max) [best, max] = [f, n];
  return best;
}

/** Định dạng quan sát được của một đoạn. */
export function observeFormat(p: ParaFmt, doc: Pick<DocFormat, 'defaultSizePt' | 'defaultFont'>): RoleFormat {
  const fmt: RoleFormat = {
    align: p.align,
    sizePt: round(dominantSize(p, doc.defaultSizePt ?? 11), 0.5),
    bold: majority(p, (r) => r.bold),
    italic: majority(p, (r) => r.italic),
    caps: majority(p, (r) => r.caps) || (isAllCaps(p.text) && /\p{L}{2,}/u.test(p.text)),
    firstLineMm: round(p.firstLineMm, 0.5),
    leftMm: round(p.leftMm, 0.5),
    spaceBeforePt: round(p.spaceBeforePt, 0.5),
    spaceAfterPt: round(p.spaceAfterPt, 0.5),
    lineSpacing: round(p.lineSpacing, 0.05),
  };
  const font = dominantFont(p, doc.defaultFont);
  if (font) fmt.font = font;
  return fmt;
}

/* ------------------------------------------------------------------ */
/* Học vai trò                                                         */
/* ------------------------------------------------------------------ */

const ctxKey = (p: ParaFmt): string => (p.table ? `T${p.table.col}${p.table.row === 0 ? 'h' : ''}` : 'B');
function signature(f: RoleFormat, p: ParaFmt): string {
  return [
    ctxKey(p),
    f.align,
    f.sizePt,
    f.bold ? 'b' : '-',
    f.italic ? 'i' : '-',
    f.caps ? 'C' : '-',
    round(f.firstLineMm, 3),
    round(f.leftMm, 3),
    round(f.spaceBeforePt, 3),
  ].join('|');
}

function distance(a: RoleFormat, ctxA: string, b: RoleFormat, ctxB: string): number {
  return (
    Math.abs(a.sizePt - b.sizePt) +
    (a.bold === b.bold ? 0 : 3) +
    (a.italic === b.italic ? 0 : 2) +
    (a.align === b.align ? 0 : 2) +
    (ctxA === ctxB ? 0 : 4) +
    Math.abs(a.firstLineMm - b.firstLineMm) / 5 +
    Math.abs(a.leftMm - b.leftMm) / 5
  );
}

interface Cluster {
  fmt: RoleFormat;
  ctx: string;
  paras: ParaFmt[];
}

const mean = (xs: number[]): number => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0);
const mode = <T>(xs: T[]): T | undefined => {
  const m = new Map<T, number>();
  for (const x of xs) m.set(x, (m.get(x) ?? 0) + 1);
  let best: T | undefined;
  let max = 0;
  for (const [k, n] of m) if (n > max) [best, max] = [k, n];
  return best;
};

function describe(c: { features: RoleFeatures; inTable: number; tableCol?: number; format: RoleFormat }, idx: number): string {
  const f = c.features;
  let base: string;
  if (f.roman > 0.5) base = 'Mục lớn (I., II.)';
  else if (f.decimal > 0.5) base = 'Mục (1., 2.)';
  else if (f.letter > 0.5) base = 'Ý nhỏ (a), b))';
  else if (f.bullet > 0.5 || f.listed > 0.5) base = 'Gạch đầu dòng / danh sách';
  else if (f.caps > 0.6) base = 'Tiêu đề viết hoa';
  else if (f.colon > 0.5) base = 'Nhãn (kết thúc bằng dấu :)';
  else if (c.format.bold && f.avgWords <= 12) base = 'Tiêu đề đậm';
  else base = 'Nội dung';
  const where = c.inTable > 0.5 ? ` trong bảng, cột ${(c.tableCol ?? 0) + 1}` : '';
  return `${base}${where} (${idx})`;
}

export function learnProfile(doc: DocFormat, name = 'Mẫu của tôi', createdFrom?: string): FormatProfile {
  const paras = doc.paragraphs.filter((p) => p.text.trim().length > 0);
  if (paras.length === 0) throw new Error('File mẫu không có nội dung chữ để học.');

  // 1. Gom theo chữ ký định dạng.
  const clusters = new Map<string, Cluster>();
  for (const p of paras) {
    const fmt = observeFormat(p, doc);
    const key = signature(fmt, p);
    const c = clusters.get(key) ?? { fmt, ctx: ctxKey(p), paras: [] };
    c.paras.push(p);
    clusters.set(key, c);
  }

  // 2. Cụm chỉ có một đoạn: gộp vào cụm gần nhất nếu định dạng gần giống (sai lệch nhỏ),
  //    còn nếu khác hẳn (ví dụ tiêu đề văn bản chỉ xuất hiện một lần) thì giữ làm một vai trò riêng.
  let list = [...clusters.values()].sort((a, b) => b.paras.length - a.paras.length);
  const big = list.filter((c) => c.paras.length >= 2);
  const lone = list.filter((c) => c.paras.length < 2);
  if (big.length > 0) {
    const keptLone: Cluster[] = [];
    for (const s of lone) {
      let best = big[0] as Cluster;
      let d = Infinity;
      for (const k of big) {
        const dd = distance(s.fmt, s.ctx, k.fmt, k.ctx);
        if (dd < d) [best, d] = [k, dd];
      }
      if (d < 3) best.paras.push(...s.paras);
      else keptLone.push(s);
    }
    list = [...big, ...keptLone];
  }

  // 3. Dựng vai trò.
  const total = paras.length;
  const roles: Role[] = list.map((c, i) => {
    const fs = c.paras.map(featuresOf);
    const rate = (pick: (f: ParaFeatures) => boolean): number => fs.filter(pick).length / fs.length;
    const features: RoleFeatures = {
      roman: rate((f) => f.roman),
      decimal: rate((f) => f.decimal),
      letter: rate((f) => f.letter),
      bullet: rate((f) => f.bullet),
      caps: rate((f) => f.caps),
      colon: rate((f) => f.colon),
      listed: rate((f) => f.listed),
      avgWords: Math.round(mean(fs.map((f) => f.words)) * 10) / 10,
    };
    const inTable = c.paras.filter((p) => p.table).length / c.paras.length;
    const firstRow = c.paras.filter((p) => p.table?.row === 0).length / c.paras.length;
    const col = mode(c.paras.filter((p) => p.table).map((p) => (p.table as { col: number }).col));
    // Định dạng đại diện: lấy giá trị phổ biến nhất của cụm (các đoạn gộp vào có thể khác nhau chút ít).
    const f0 = c.fmt;
    const format: RoleFormat = {
      ...f0,
      sizePt: mode(c.paras.map((p) => observeFormat(p, doc).sizePt)) ?? f0.sizePt,
    };
    const role: Role = {
      id: `r${i + 1}`,
      name: '',
      count: c.paras.length,
      format,
      inTable,
      firstRow,
      features,
      examples: c.paras.slice(0, 3).map((p) => p.text.trim().slice(0, 60)),
    };
    if (col !== undefined && inTable > 0.5) role.tableCol = col;
    role.name = describe(role, i + 1);
    return role;
  });

  const profile: FormatProfile = { version: 1, name, roles };
  if (createdFrom) profile.createdFrom = createdFrom;
  const font = mode(paras.map((p) => dominantFont(p, doc.defaultFont)).filter((x): x is string => !!x));
  if (font) profile.font = font;
  if (doc.page) profile.page = pageSpec(doc.page);
  void total;
  return profile;
}

function pageSpec(p: PageFmt): PageSpec {
  return {
    widthMm: round(p.widthMm, 1),
    heightMm: round(p.heightMm, 1),
    marginTopMm: round(p.marginTopMm, 0.5),
    marginBottomMm: round(p.marginBottomMm, 0.5),
    marginLeftMm: round(p.marginLeftMm, 0.5),
    marginRightMm: round(p.marginRightMm, 0.5),
  };
}

/* ------------------------------------------------------------------ */
/* Đoán vai trò cho một đoạn                                            */
/* ------------------------------------------------------------------ */

const EPS = 0.05;
const ll = (rate: number, x: boolean): number => Math.log(x ? rate + EPS : 1 - rate + EPS);

export function roleFor(profile: FormatProfile, p: ParaFmt): Role {
  const f = featuresOf(p);
  const total = profile.roles.reduce((s, r) => s + r.count, 0);
  let best = profile.roles[0] as Role;
  let bestScore = -Infinity;
  for (const r of profile.roles) {
    let s = 0.3 * Math.log(r.count / total);
    s += ll(r.features.roman, f.roman);
    s += ll(r.features.decimal, f.decimal);
    s += ll(r.features.letter, f.letter);
    s += ll(r.features.bullet, f.bullet);
    s += ll(r.features.caps, f.caps);
    s += ll(r.features.colon, f.colon);
    s += ll(r.features.listed, f.listed);
    s += ll(r.inTable, !!p.table) * 1.5;
    s += ll(r.firstRow, p.table?.row === 0) * 1.5;
    if (p.table && r.tableCol !== undefined && r.tableCol !== p.table.col) s -= 1.5;
    s -= 0.5 * Math.abs(Math.log((f.words + 1) / (r.features.avgWords + 1)));
    if (s > bestScore) [best, bestScore] = [r, s];
  }
  return best;
}

/* ------------------------------------------------------------------ */
/* So sánh và kế hoạch sửa                                              */
/* ------------------------------------------------------------------ */

export type Field = keyof RoleFormat;

export const FIELD_NAMES: Record<Field, string> = {
  align: 'căn lề',
  sizePt: 'cỡ chữ',
  bold: 'chữ đậm',
  italic: 'chữ nghiêng',
  caps: 'viết hoa toàn bộ',
  font: 'font chữ',
  firstLineMm: 'thụt đầu dòng',
  leftMm: 'thụt lề trái',
  spaceBeforePt: 'khoảng cách trước đoạn',
  spaceAfterPt: 'khoảng cách sau đoạn',
  lineSpacing: 'giãn dòng',
};

const ALIGN_NAMES: Record<Align, string> = { left: 'trái', center: 'giữa', right: 'phải', justify: 'đều hai bên' };

export function showValue(field: Field, v: unknown): string {
  if (typeof v === 'boolean') return v ? 'có' : 'không';
  if (field === 'align') return ALIGN_NAMES[v as Align] ?? String(v);
  if (field === 'sizePt' || field === 'spaceBeforePt' || field === 'spaceAfterPt') return `${v} pt`;
  if (field === 'firstLineMm' || field === 'leftMm') return `${v} mm`;
  if (field === 'lineSpacing') return `×${v}`;
  return String(v);
}

const TOL: Partial<Record<Field, number>> = {
  sizePt: 0.5,
  firstLineMm: 1.5,
  leftMm: 1.5,
  spaceBeforePt: 1.5,
  spaceAfterPt: 1.5,
  lineSpacing: 0.08,
};

export interface FieldDiff {
  field: Field;
  expected: RoleFormat[Field];
  actual: RoleFormat[Field];
}

export function diffFormat(expected: RoleFormat, actual: RoleFormat, profileFont?: string): FieldDiff[] {
  const out: FieldDiff[] = [];
  const fields: Field[] = ['align', 'sizePt', 'bold', 'italic', 'caps', 'firstLineMm', 'leftMm', 'spaceBeforePt', 'spaceAfterPt', 'lineSpacing'];
  for (const f of fields) {
    const e = expected[f];
    const a = actual[f];
    const tol = TOL[f];
    const same = tol !== undefined ? Math.abs((e as number) - (a as number)) <= tol : e === a;
    if (!same) out.push({ field: f, expected: e, actual: a });
  }
  const ef = expected.font ?? profileFont;
  if (ef && actual.font && ef.toLowerCase() !== actual.font.toLowerCase()) out.push({ field: 'font', expected: ef, actual: actual.font });
  return out;
}

export interface ParagraphIssue {
  index: number;
  text: string;
  roleId: string;
  roleName: string;
  diffs: FieldDiff[];
}

export interface PageIssue {
  field: string;
  expected: string;
  actual: string;
}

export interface CheckResult {
  paragraphs: number;
  issues: ParagraphIssue[];
  page: PageIssue[];
  /** Vai trò đã đoán cho từng đoạn không trống (theo chỉ số đoạn) */
  roleOf: Map<number, string>;
}

export function checkAgainstProfile(profile: FormatProfile, doc: DocFormat): CheckResult {
  const issues: ParagraphIssue[] = [];
  const roleOf = new Map<number, string>();
  let checked = 0;
  for (const p of doc.paragraphs) {
    if (!p.text.trim()) continue;
    checked++;
    const role = roleFor(profile, p);
    roleOf.set(p.index, role.id);
    const diffs = diffFormat(role.format, observeFormat(p, doc), profile.font);
    if (diffs.length) issues.push({ index: p.index, text: p.text.trim().slice(0, 70), roleId: role.id, roleName: role.name, diffs });
  }
  const page: PageIssue[] = [];
  if (profile.page && doc.page) {
    const a = pageSpec(doc.page);
    const e = profile.page;
    const cmp = (name: string, ev: number, av: number, tol: number): void => {
      if (Math.abs(ev - av) > tol) page.push({ field: name, expected: `${ev} mm`, actual: `${av} mm` });
    };
    cmp('khổ giấy (rộng)', e.widthMm, a.widthMm, 1.5);
    cmp('khổ giấy (cao)', e.heightMm, a.heightMm, 1.5);
    cmp('lề trên', e.marginTopMm, a.marginTopMm, 0.7);
    cmp('lề dưới', e.marginBottomMm, a.marginBottomMm, 0.7);
    cmp('lề trái', e.marginLeftMm, a.marginLeftMm, 0.7);
    cmp('lề phải', e.marginRightMm, a.marginRightMm, 0.7);
  }
  return { paragraphs: checked, issues, page, roleOf };
}

/** Kế hoạch sửa: với mỗi đoạn lệch, các thuộc tính cần đặt lại. */
export interface FixItem {
  index: number;
  text: string;
  roleId: string;
  set: Partial<RoleFormat>;
}

export function planFixes(profile: FormatProfile, doc: DocFormat): FixItem[] {
  const result = checkAgainstProfile(profile, doc);
  return result.issues.map((i) => {
    const set: Partial<RoleFormat> = {};
    for (const d of i.diffs) (set as Record<string, unknown>)[d.field] = d.expected;
    return { index: i.index, text: i.text, roleId: i.roleId, set };
  });
}
