/**
 * Đọc định dạng của một file Word (.docx hoặc OOXML do Word.getOoxml() trả về).
 * Chỉ đọc những gì công cụ cần: lề trang, từng đoạn (căn lề, thụt lề, giãn dòng, chữ đậm/nghiêng, cỡ chữ, font)
 * và độ rộng cột của các bảng. Có xử lý kế thừa Style cơ bản (docDefaults, basedOn, pStyle).
 */
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const TWIP_PER_MM = 1440 / 25.4;

export interface RunFmt {
  text: string;
  bold: boolean;
  italic: boolean;
  caps: boolean;
  sizePt?: number;
  font?: string;
}

export interface ParaFmt {
  /** Thứ tự trong tài liệu, tính cả đoạn trống và đoạn trong bảng. */
  index: number;
  text: string;
  align: 'left' | 'center' | 'right' | 'justify';
  firstLineMm: number;
  leftMm: number;
  rightMm: number;
  spaceBeforePt: number;
  spaceAfterPt: number;
  /** Bội số giãn dòng (1 = đơn). Với "exact/atLeast" là giá trị ước lượng. */
  lineSpacing: number;
  runs: RunFmt[];
  /** Mức của danh sách tự động đánh số/gạch đầu dòng của Word (0 là mức đầu); undefined nếu không phải danh sách. */
  listLevel?: number;
  table?: { id: number; row: number; col: number };
}

export interface PageFmt {
  widthMm: number;
  heightMm: number;
  marginTopMm: number;
  marginBottomMm: number;
  marginLeftMm: number;
  marginRightMm: number;
  titlePg: boolean;
}

export interface BorderFmt {
  sizePt: number;
  /** hex không có # (chữ hoa) */
  color: string;
  style: 'single' | 'double' | 'dashed' | 'dotted';
}
export type BorderSide = 'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV';
export type TableBorders = Record<BorderSide, BorderFmt | null>;

/** Khung của một bảng (chỉ bảng ở mức cao nhất, không tính bảng lồng nhau). */
export interface TableFmt {
  id: number;
  rows: number;
  colWidthsMm: number[];
  align: 'left' | 'center' | 'right';
  borders: TableBorders;
  /** Màu nền hàng đầu (hex), null nếu không tô. */
  headerShading: string | null;
  headerRepeat: boolean;
  vAlign: 'top' | 'center' | 'bottom';
  /** [trên, phải, dưới, trái] mm; null = mặc định của Word. */
  cellMarginMm: [number, number, number, number] | null;
}

export interface DocFormat {
  paragraphs: ParaFmt[];
  /** undefined nếu nguồn không có thông tin trang (ví dụ OOXML của một vùng chọn). */
  page?: PageFmt;
  tables: TableFmt[];
  defaultFont?: string;
  defaultSizePt?: number;
  /** undefined nếu nguồn không có phần đầu trang để xét. */
  hasPageNumber?: boolean;
}

/* ------------------------------------------------------------------ */
/* Bộ đọc XML: trình duyệt dùng DOMParser; Node (test, CLI) tự cấp.     */
/* ------------------------------------------------------------------ */

type XmlParser = (xml: string) => Document;
let customParser: XmlParser | null = null;
export function setXmlParser(p: XmlParser): void {
  customParser = p;
}
function parseXml(xml: string): Document {
  if (customParser) return customParser(xml);
  return new DOMParser().parseFromString(xml, 'application/xml');
}

const kids = (el: Element, name: string): Element[] =>
  Array.from(el.childNodes).filter(
    (n): n is Element => n.nodeType === 1 && (n as Element).localName === name && (n as Element).namespaceURI === W,
  );
const kid = (el: Element | null | undefined, name: string): Element | null => (el ? kids(el, name)[0] ?? null : null);
const attr = (el: Element | null | undefined, name: string): string | null => {
  if (!el) return null;
  return el.getAttributeNS(W, name) ?? el.getAttribute(`w:${name}`);
};
const num = (v: string | null): number | undefined => {
  if (v === null || v === '') return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
};
/** Phần tử bật/tắt kiểu <w:b/> hoặc <w:b w:val="0"/> */
function toggle(parent: Element | null, name: string): boolean | undefined {
  const el = kid(parent, name);
  if (!el) return undefined;
  const v = attr(el, 'val');
  return !(v === '0' || v === 'false' || v === 'off');
}
const mm = (twips: number | undefined): number | undefined => (twips === undefined ? undefined : twips / TWIP_PER_MM);

/* ------------------------------------------------------------------ */
/* Thuộc tính đoạn / chữ, có kế thừa                                    */
/* ------------------------------------------------------------------ */

interface PProps {
  align?: ParaFmt['align'];
  firstLine?: number;
  left?: number;
  right?: number;
  before?: number;
  after?: number;
  line?: number;
  lineRule?: string;
  listLevel?: number;
}
interface RProps {
  bold?: boolean;
  italic?: boolean;
  caps?: boolean;
  size?: number;
  font?: string;
}

function readPPr(pPr: Element | null): PProps {
  const p: PProps = {};
  if (!pPr) return p;
  const jc = attr(kid(pPr, 'jc'), 'val');
  if (jc) p.align = jc === 'both' || jc === 'distribute' ? 'justify' : jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : 'left';
  const ind = kid(pPr, 'ind');
  if (ind) {
    const first = num(attr(ind, 'firstLine'));
    const hang = num(attr(ind, 'hanging'));
    if (first !== undefined) p.firstLine = first;
    else if (hang !== undefined) p.firstLine = -hang;
    const left = num(attr(ind, 'left') ?? attr(ind, 'start'));
    const right = num(attr(ind, 'right') ?? attr(ind, 'end'));
    if (left !== undefined) p.left = left;
    if (right !== undefined) p.right = right;
  }
  const numPr = kid(pPr, 'numPr');
  if (numPr) {
    const numId = num(attr(kid(numPr, 'numId'), 'val'));
    if (numId !== undefined && numId !== 0) p.listLevel = num(attr(kid(numPr, 'ilvl'), 'val')) ?? 0;
    else if (numId === 0) p.listLevel = -1; // tắt danh sách kế thừa từ Style
  }
  const sp = kid(pPr, 'spacing');
  if (sp) {
    const before = num(attr(sp, 'before'));
    const after = num(attr(sp, 'after'));
    const line = num(attr(sp, 'line'));
    if (before !== undefined) p.before = before;
    if (after !== undefined) p.after = after;
    if (line !== undefined) {
      p.line = line;
      p.lineRule = attr(sp, 'lineRule') ?? 'auto';
    }
  }
  return p;
}

function readRPr(rPr: Element | null): RProps {
  const r: RProps = {};
  if (!rPr) return r;
  const b = toggle(rPr, 'b');
  const i = toggle(rPr, 'i');
  const caps = toggle(rPr, 'caps');
  if (b !== undefined) r.bold = b;
  if (i !== undefined) r.italic = i;
  if (caps !== undefined) r.caps = caps;
  const sz = num(attr(kid(rPr, 'sz'), 'val'));
  if (sz !== undefined) r.size = sz / 2;
  const fonts = kid(rPr, 'rFonts');
  const font = attr(fonts, 'ascii') ?? attr(fonts, 'hAnsi');
  if (font) r.font = font;
  return r;
}

const BORDER_SIDES: [BorderSide, string[]][] = [
  ['top', ['top']],
  ['bottom', ['bottom']],
  ['left', ['left', 'start']],
  ['right', ['right', 'end']],
  ['insideH', ['insideH']],
  ['insideV', ['insideV']],
];

function readBorder(el: Element | null): BorderFmt | null | undefined {
  if (!el) return undefined;
  const val = attr(el, 'val');
  if (!val || val === 'nil' || val === 'none') return null;
  const color = attr(el, 'color');
  const style = val === 'double' ? 'double' : val === 'dashed' || val === 'dashSmallGap' ? 'dashed' : val === 'dotted' ? 'dotted' : 'single';
  return {
    sizePt: (num(attr(el, 'sz')) ?? 4) / 8,
    color: !color || color === 'auto' ? '000000' : color.toUpperCase(),
    style,
  };
}

function readBorders(tblBorders: Element | null): Partial<TableBorders> {
  const out: Partial<TableBorders> = {};
  if (!tblBorders) return out;
  for (const [side, names] of BORDER_SIDES) {
    for (const n of names) {
      const b = readBorder(kid(tblBorders, n));
      if (b !== undefined) {
        out[side] = b;
        break;
      }
    }
  }
  return out;
}

function readCellMar(el: Element | null): [number, number, number, number] | null {
  if (!el) return null;
  const side = (names: string[]): number => {
    for (const n of names) {
      const w = num(attr(kid(el, n), 'w'));
      if (w !== undefined) return w / TWIP_PER_MM;
    }
    return 0;
  };
  return [side(['top']), side(['right', 'end']), side(['bottom']), side(['left', 'start'])];
}

interface TableStyleDef {
  basedOn?: string;
  borders: Partial<TableBorders>;
  cellMar: [number, number, number, number] | null;
}

const assignDefined = <T extends object>(target: T, src: Partial<T>): T => {
  for (const [k, v] of Object.entries(src)) if (v !== undefined) (target as Record<string, unknown>)[k] = v;
  return target;
};

interface StyleDef {
  basedOn?: string;
  p: PProps;
  r: RProps;
  isDefault: boolean;
}

function readStyles(stylesXml: string | undefined): {
  styles: Map<string, StyleDef>;
  tableStyles: Map<string, TableStyleDef>;
  defP: PProps;
  defR: RProps;
} {
  const styles = new Map<string, StyleDef>();
  const tableStyles = new Map<string, TableStyleDef>();
  const defP: PProps = {};
  const defR: RProps = {};
  if (!stylesXml) return { styles, tableStyles, defP, defR };
  const root = parseXml(stylesXml).documentElement;
  const dd = kid(root, 'docDefaults');
  assignDefined(defR, readRPr(kid(kid(dd, 'rPrDefault'), 'rPr')));
  assignDefined(defP, readPPr(kid(kid(dd, 'pPrDefault'), 'pPr')));
  for (const s of kids(root, 'style')) {
    if (attr(s, 'type') === 'table') {
      const id = attr(s, 'styleId');
      if (id) {
        const tblPr = kid(s, 'tblPr');
        const basedOn = attr(kid(s, 'basedOn'), 'val') ?? undefined;
        tableStyles.set(id, {
          ...(basedOn ? { basedOn } : {}),
          borders: readBorders(kid(tblPr, 'tblBorders')),
          cellMar: readCellMar(kid(tblPr, 'tblCellMar')),
        });
      }
      continue;
    }
    if (attr(s, 'type') !== 'paragraph') continue;
    const id = attr(s, 'styleId');
    if (!id) continue;
    const basedOn = attr(kid(s, 'basedOn'), 'val') ?? undefined;
    styles.set(id, {
      ...(basedOn ? { basedOn } : {}),
      p: readPPr(kid(s, 'pPr')),
      r: readRPr(kid(s, 'rPr')),
      isDefault: attr(s, 'default') === '1',
    });
  }
  return { styles, tableStyles, defP, defR };
}

function styleChain(styles: Map<string, StyleDef>, id: string | undefined): StyleDef[] {
  const chain: StyleDef[] = [];
  const seen = new Set<string>();
  let cur = id;
  while (cur && styles.has(cur) && !seen.has(cur)) {
    seen.add(cur);
    const s = styles.get(cur) as StyleDef;
    chain.unshift(s);
    cur = s.basedOn;
  }
  return chain;
}

/* ------------------------------------------------------------------ */
/* Đọc document.xml                                                     */
/* ------------------------------------------------------------------ */

export function parseDocumentXml(
  documentXml: string,
  extra: { stylesXml?: string; headerXmls?: string[] } = {},
): DocFormat {
  const doc = parseXml(documentXml).documentElement;
  const body = kid(doc, 'body');
  if (!body) throw new Error('Không tìm thấy nội dung văn bản (w:body).');
  const { styles, tableStyles, defP, defR } = readStyles(extra.stylesXml);
  const defaultStyleId = [...styles.entries()].find(([, s]) => s.isDefault)?.[0];

  const paragraphs: ParaFmt[] = [];
  const tables: DocFormat['tables'] = [];
  let sectPr: Element | null = null;

  const readParagraph = (p: Element, table?: ParaFmt['table']): void => {
    const pPrEl = kid(p, 'pPr');
    const styleId = attr(kid(pPrEl, 'pStyle'), 'val') ?? defaultStyleId ?? undefined;
    const chain = styleChain(styles, styleId);
    const P: PProps = { ...defP };
    const R: RProps = { ...defR };
    for (const s of chain) {
      assignDefined(P, s.p);
      assignDefined(R, s.r);
    }
    assignDefined(P, readPPr(pPrEl));
    // Đoạn này có thể mang sectPr (kết thúc một phần): lấy làm thông tin trang.
    const inlineSect = kid(pPrEl, 'sectPr');
    if (inlineSect) sectPr = inlineSect;
    // Dấu đoạn có thể mang định dạng chữ mặc định cho cả đoạn
    assignDefined(R, readRPr(kid(pPrEl, 'rPr')));

    const runs: RunFmt[] = [];
    for (const r of Array.from(p.getElementsByTagNameNS(W, 'r'))) {
      let text = '';
      for (const n of Array.from(r.childNodes)) {
        const el = n as Element;
        if (el.nodeType !== 1 || el.namespaceURI !== W) continue;
        if (el.localName === 't') text += el.textContent ?? '';
        else if (el.localName === 'tab') text += '\t';
        else if (el.localName === 'br') text += '\n';
      }
      if (!text) continue;
      const rp: RProps = { ...R };
      assignDefined(rp, readRPr(kid(r, 'rPr')));
      const run: RunFmt = { text, bold: !!rp.bold, italic: !!rp.italic, caps: !!rp.caps };
      if (rp.size !== undefined) run.sizePt = rp.size;
      if (rp.font) run.font = rp.font;
      runs.push(run);
    }

    const size = P.line === undefined ? 1 : P.lineRule === 'auto' ? P.line / 240 : P.line / 20 / ((R.size ?? 12) * 1.15);
    const fmt: ParaFmt = {
      index: paragraphs.length,
      text: runs.map((r) => r.text).join(''),
      align: P.align ?? 'left',
      firstLineMm: mm(P.firstLine) ?? 0,
      leftMm: mm(P.left) ?? 0,
      rightMm: mm(P.right) ?? 0,
      spaceBeforePt: (P.before ?? 0) / 20,
      spaceAfterPt: (P.after ?? 0) / 20,
      lineSpacing: Math.round(size * 100) / 100,
      runs,
    };
    if (P.listLevel !== undefined && P.listLevel >= 0) fmt.listLevel = P.listLevel;
    if (table) fmt.table = table;
    paragraphs.push(fmt);
  };

  for (const node of Array.from(body.childNodes)) {
    const el = node as Element;
    if (el.nodeType !== 1 || el.namespaceURI !== W) continue;
    if (el.localName === 'p') readParagraph(el);
    else if (el.localName === 'sectPr') sectPr = el;
    else if (el.localName === 'tbl') {
      const id = tables.length;
      const grid = kid(el, 'tblGrid');
      const tblPr = kid(el, 'tblPr');
      // Viền và lề ô: Style bảng (có kế thừa) rồi đến thuộc tính trực tiếp của bảng.
      const borders: TableBorders = { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
      let cellMar: [number, number, number, number] | null = null;
      const styleChain: TableStyleDef[] = [];
      const seenStyles = new Set<string>();
      for (let sid = attr(kid(tblPr, 'tblStyle'), 'val') ?? undefined; sid && tableStyles.has(sid) && !seenStyles.has(sid); ) {
        seenStyles.add(sid);
        const st = tableStyles.get(sid) as TableStyleDef;
        styleChain.unshift(st);
        sid = st.basedOn;
      }
      for (const st of styleChain) {
        Object.assign(borders, st.borders);
        if (st.cellMar) cellMar = st.cellMar;
      }
      Object.assign(borders, readBorders(kid(tblPr, 'tblBorders')));
      cellMar = readCellMar(kid(tblPr, 'tblCellMar')) ?? cellMar;
      const jc = attr(kid(tblPr, 'jc'), 'val');

      const rows = kids(el, 'tr');
      const firstRow = rows[0];
      const headerRepeat = !!kid(kid(firstRow, 'trPr'), 'tblHeader');
      const shadeVals: string[] = [];
      const vAligns: string[] = [];
      rows.forEach((tr, rowIdx) => {
        kids(tr, 'tc').forEach((tc, col) => {
          const tcPr = kid(tc, 'tcPr');
          const v = attr(kid(tcPr, 'vAlign'), 'val');
          if (v) vAligns.push(v);
          const fill = attr(kid(tcPr, 'shd'), 'fill');
          if (rowIdx === 0 && fill && fill !== 'auto') shadeVals.push(fill.toUpperCase());
          for (const p of kids(tc, 'p')) readParagraph(p, { id, row: rowIdx, col });
        });
      });
      const vMode = ((): 'top' | 'center' | 'bottom' => {
        const m = new Map<string, number>();
        for (const v of vAligns) m.set(v, (m.get(v) ?? 0) + 1);
        const top = [...m.entries()].sort((x, y) => y[1] - x[1])[0]?.[0];
        return top === 'center' ? 'center' : top === 'bottom' ? 'bottom' : 'top';
      })();
      tables.push({
        id,
        rows: rows.length,
        colWidthsMm: kids(grid as Element, 'gridCol').map((g) => mm(num(attr(g, 'w'))) ?? 0),
        align: jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : 'left',
        borders,
        headerShading: shadeVals[0] ?? null,
        headerRepeat,
        vAlign: vMode,
        cellMarginMm: cellMar,
      });
    }
  }

  const out: DocFormat = { paragraphs, tables };
  const sp = sectPr as Element | null;
  const pgSz = kid(sp, 'pgSz');
  const pgMar = kid(sp, 'pgMar');
  if (pgSz && pgMar) {
    out.page = {
      widthMm: (mm(num(attr(pgSz, 'w'))) ?? 0),
      heightMm: (mm(num(attr(pgSz, 'h'))) ?? 0),
      marginTopMm: mm(num(attr(pgMar, 'top'))) ?? 0,
      marginBottomMm: mm(num(attr(pgMar, 'bottom'))) ?? 0,
      marginLeftMm: mm(num(attr(pgMar, 'left'))) ?? 0,
      marginRightMm: mm(num(attr(pgMar, 'right'))) ?? 0,
      titlePg: !!kid(sp, 'titlePg'),
    };
  }
  if (defR.font) out.defaultFont = defR.font;
  if (defR.size !== undefined) out.defaultSizePt = defR.size;
  if (extra.headerXmls) out.hasPageNumber = extra.headerXmls.some((h) => /instr="[^"]*\bPAGE\b|<w:instrText[^>]*>[^<]*\bPAGE\b/.test(h));
  return out;
}

/** Đọc file .docx (byte). */
export async function readDocx(bytes: Uint8Array | ArrayBuffer): Promise<DocFormat> {
  const zip = await JSZip.loadAsync(bytes);
  const documentXml = await zip.file('word/document.xml')?.async('string');
  if (!documentXml) throw new Error('File không phải .docx hợp lệ (thiếu word/document.xml).');
  const stylesXml = await zip.file('word/styles.xml')?.async('string');
  const headerXmls = await Promise.all(
    Object.keys(zip.files)
      .filter((n) => /^word\/header\d*\.xml$/.test(n))
      .map((n) => (zip.file(n) as JSZip.JSZipObject).async('string')),
  );
  return parseDocumentXml(documentXml, { ...(stylesXml ? { stylesXml } : {}), headerXmls });
}

/** Đọc "Flat OPC" mà Word.Body.getOoxml() trả về (một file XML chứa nhiều phần của gói .docx). */
export function readFlatOpc(flatXml: string): DocFormat {
  const root = parseXml(flatXml).documentElement;
  const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
  const parts = new Map<string, Element>();
  for (const part of Array.from(root.getElementsByTagNameNS(PKG, 'part'))) {
    const name = part.getAttributeNS(PKG, 'name') ?? part.getAttribute('pkg:name');
    const data = Array.from(part.childNodes).find((n) => n.nodeType === 1 && (n as Element).localName === 'xmlData') as Element | undefined;
    const content = data ? (Array.from(data.childNodes).find((n) => n.nodeType === 1) as Element | undefined) : undefined;
    if (name && content) parts.set(name, content);
  }
  const serialize = (el: Element | undefined): string | undefined => (el ? new XMLSerializerShim().serializeToString(el) : undefined);
  const documentEl = parts.get('/word/document.xml');
  if (!documentEl) throw new Error('OOXML của Word không có phần word/document.xml.');
  const headerXmls = [...parts.entries()].filter(([n]) => /^\/word\/header\d*\.xml$/.test(n)).map(([, el]) => serialize(el) as string);
  const stylesXml = serialize(parts.get('/word/styles.xml'));
  return parseDocumentXml(serialize(documentEl) as string, { ...(stylesXml ? { stylesXml } : {}), headerXmls });
}

/** Trình duyệt có XMLSerializer sẵn; Node dùng bản của xmldom do bên gọi cấp qua setXmlSerializer. */
let customSerializer: ((el: Element) => string) | null = null;
export function setXmlSerializer(fn: (el: Element) => string): void {
  customSerializer = fn;
}
class XMLSerializerShim {
  serializeToString(el: Element): string {
    return customSerializer ? customSerializer(el) : new XMLSerializer().serializeToString(el);
  }
}
