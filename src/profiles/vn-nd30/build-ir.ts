import type {
  Align,
  BlockNode,
  DocumentIR,
  ParagraphNode,
  Run,
  TableCellNode,
} from '../../ir';
import { estimateWidthMm, toUpperVi, wordCount } from '../../text';
import layoutJson from './layout.json';
import { splitAgencyName } from './classify';
import type { BodyBlock, Model, PlaceDate } from './types';

/* Kiểu lỏng của layout.json, để truy cập các khoá tuỳ chọn mà không bị TS suy ra union hẹp. */
export interface Style {
  sizePt?: number;
  sizeRangePt?: [number, number];
  bold?: boolean;
  italic?: boolean;
  caps?: boolean;
  align?: Align;
  spaceBeforePt?: number;
  ruleRatio?: number;
  ruleMatchesText?: boolean;
}
export interface Layout {
  font: string;
  page: {
    widthMm: number;
    heightMm: number;
    margins: {
      topMm: number;
      bottomMm: number;
      leftMm: number;
      rightMm: number;
      range?: { topMm: [number, number]; bottomMm: [number, number]; leftMm: [number, number]; rightMm: [number, number] };
    };
    headerDistanceMm: number;
    footerDistanceMm: number;
    pageNumber: { sizePt: number; hideOnFirstPage: boolean } | null;
  };
  defaults: {
    bodySizePt: number;
    bodySizeRangePt?: [number, number];
    lineSpacing: number;
    lineSpacingRange?: [number, number];
    paragraphSpaceBeforePt: number;
    firstLineIndentMm: number;
    firstLineIndentRangeMm?: [number, number];
  };
  header: {
    leftColumnMm: number;
    rightColumnMm: number;
    parentOrg: Style;
    issuer: Style;
    nationalTitle: Style;
    motto: Style;
    number: Style;
    subjectCongVan: Style;
    placeDate: Style;
  };
  title: { docTypeName: Style; subject: Style };
  addressee: Style;
  body: {
    paragraph: Style;
    basis: { italicFor: string[] };
    enact: Style;
    note: Style;
    authority: Style;
    chapter: Style;
    article: { shortTitleMaxWords: number };
    headings: Record<string, Style>;
  };
  closing: {
    leftColumnMm: number;
    rightColumnMm: number;
    recipientsTitle: Style;
    recipientItem: Style;
    signer: Style;
    signatureSpacePt: number;
    spaceBeforePt: number;
  };
}

export const ND30_LAYOUT = layoutJson as unknown as Layout;

const NBSP_ELLIPSIS = '…';

/** Chọn cỡ chữ lớn nhất trong khoảng cho phép mà vừa một dòng; không vừa thì dùng cỡ nhỏ nhất. */
function fitSize(text: string, st: Style, widthMm: number): number {
  const preferred = st.sizePt ?? 13;
  const min = st.sizeRangePt?.[0] ?? preferred;
  for (let s = preferred; s >= min; s -= 0.5) {
    if (estimateWidthMm(text, s, st.bold) <= widthMm - 1) return s;
  }
  return min;
}

function para(text: string, st: Style, extra: Partial<ParagraphNode> = {}): ParagraphNode {
  const run: Run = { text: st.caps ? toUpperVi(text) : text };
  if (st.bold) run.bold = true;
  if (st.italic) run.italic = true;
  const node: ParagraphNode = {
    type: 'paragraph',
    runs: [run],
    align: st.align ?? 'left',
    ...extra,
  };
  if (st.sizePt !== undefined && node.sizePt === undefined) node.sizePt = st.sizePt;
  if (st.spaceBeforePt !== undefined && node.spaceBeforePt === undefined)
    node.spaceBeforePt = st.spaceBeforePt;
  return node;
}

export function formatDate(pd: PlaceDate): string {
  const day =
    pd.day === undefined ? NBSP_ELLIPSIS : pd.day < 10 ? `0${pd.day}` : String(pd.day);
  const month =
    pd.month === undefined ? NBSP_ELLIPSIS : pd.month <= 2 ? `0${pd.month}` : String(pd.month);
  const year = pd.year === undefined ? NBSP_ELLIPSIS : String(pd.year);
  const place = pd.place ? `${pd.place}, ` : `${NBSP_ELLIPSIS}, `;
  return `${place}ngày ${day} tháng ${month} năm ${year}`;
}

/** Danh sách có dấu gạch đầu dòng: các mục kết thúc bằng ";" và mục cuối bằng ".". */
function punctuateList(items: string[]): string[] {
  return items.map((it, i) => {
    const core = it.replace(/[;,.\s]+$/g, '');
    return `- ${core}${i === items.length - 1 ? '.' : ';'}`;
  });
}

function headerBlock(m: Model, L: Layout): BlockNode[] {
  const H = L.header;
  const left = H.leftColumnMm;
  const right = H.rightColumnMm;

  /* Hàng 1 – trái: cơ quan chủ quản + cơ quan ban hành + đường kẻ */
  const leftTop: BlockNode[] = [];
  for (const p of m.parentOrg) leftTop.push(para(p, H.parentOrg, { lineSpacing: 1 }));
  const issuerText = m.issuer ?? 'TÊN CƠ QUAN BAN HÀNH';
  const issuerLines = issuerText.split('\n').flatMap(splitAgencyName);
  for (const line of issuerLines) leftTop.push(para(line, H.issuer, { lineSpacing: 1 }));
  const widest = Math.max(
    ...issuerLines.map((l) => estimateWidthMm(toUpperVi(l), H.issuer.sizePt ?? 13, true)),
  );
  const ruleWidth = Math.max(10, Math.min(left - 4, widest * (H.issuer.ruleRatio ?? 0.4)));
  leftTop.push({ type: 'rule', widthMm: ruleWidth, spaceBeforePt: 1 });

  /* Hàng 1 – phải: quốc hiệu + tiêu ngữ + đường kẻ */
  const national = 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM';
  const motto = 'Độc lập - Tự do - Hạnh phúc';
  const nSize = fitSize(national, H.nationalTitle, right);
  const mSize = fitSize(motto, H.motto, right);
  const rightTop: BlockNode[] = [
    para(national, H.nationalTitle, { sizePt: nSize, lineSpacing: 1 }),
    para(motto, H.motto, { sizePt: mSize, lineSpacing: 1 }),
    {
      type: 'rule',
      widthMm: Math.min(right - 2, estimateWidthMm(motto, mSize, true)),
      spaceBeforePt: 1,
    },
  ];

  /* Hàng 2 – trái: số, ký hiệu (+ V/v với công văn); phải: địa danh, ngày tháng */
  const numberText = m.number ? `Số: ${m.number}` : `Số: ${NBSP_ELLIPSIS}/${NBSP_ELLIPSIS}`;
  const leftBottom: BlockNode[] = [para(numberText, H.number, { lineSpacing: 1 })];
  if (!m.docType && m.vv)
    leftBottom.push(para(`V/v ${m.vv}`, H.subjectCongVan, { lineSpacing: 1 }));

  const dateText = formatDate(m.placeDate);
  const dSize = fitSize(dateText, H.placeDate, right);
  const rightBottom: BlockNode[] = [
    para(dateText, H.placeDate, { sizePt: dSize, lineSpacing: 1 }),
  ];

  const cell = (children: BlockNode[], widthMm: number): TableCellNode => ({ children, widthMm });
  return [
    {
      type: 'table',
      rows: [
        [cell(leftTop, left), cell(rightTop, right)],
        [cell(leftBottom, left), cell(rightBottom, right)],
      ],
    },
  ];
}

function titleBlock(m: Model, L: Layout, contentMm: number): BlockNode[] {
  if (!m.docType) return [];
  const T = L.title;
  const out: BlockNode[] = [para(m.docType, T.docTypeName, { lineSpacing: 1, keepNext: true })];
  if (m.subject) {
    out.push(para(m.subject, T.subject, { lineSpacing: 1, keepNext: true }));
    const base = Math.min(estimateWidthMm(m.subject, T.subject.sizePt ?? 14, true), contentMm * 0.8);
    out.push({
      type: 'rule',
      widthMm: Math.max(15, base * (T.subject.ruleRatio ?? 0.4)),
      spaceBeforePt: 3,
    });
  }
  return out;
}

function addresseeBlock(m: Model, L: Layout): BlockNode[] {
  if (m.addressees.length === 0) return [];
  const A = L.addressee;
  if (m.addressees.length === 1) {
    return [para(`Kính gửi: ${m.addressees[0] ?? ''}`, A, { lineSpacing: L.defaults.lineSpacing })];
  }
  const items = punctuateList(m.addressees);
  return [
    para('Kính gửi:', A, { lineSpacing: L.defaults.lineSpacing, keepNext: true }),
    ...items.map((t) =>
      para(t, { ...A, spaceBeforePt: 0 }, { lineSpacing: L.defaults.lineSpacing }),
    ),
  ];
}

function bodyBlocks(m: Model, L: Layout): BlockNode[] {
  const D = L.defaults;
  const base = { lineSpacing: D.lineSpacing, spaceBeforePt: D.paragraphSpaceBeforePt };
  const italicBasis = m.docType ? L.body.basis.italicFor.includes(m.docType) : false;
  const out: BlockNode[] = [];

  const normal = (text: string, over: Partial<ParagraphNode> = {}, st: Style = {}): ParagraphNode => {
    const runText = /^[–—+•*]\s*/u.test(text) ? text.replace(/^[–—+•*]\s*/u, '- ') : text;
    const run: Run = { text: runText };
    if (st.bold) run.bold = true;
    if (st.italic) run.italic = true;
    return {
      type: 'paragraph',
      runs: [run],
      align: L.body.paragraph.align ?? 'justify',
      firstLineMm: D.firstLineIndentMm,
      ...base,
      ...over,
    };
  };

  for (const b of m.body as BodyBlock[]) {
    switch (b.kind) {
      case 'note':
        out.push(para(b.text, L.body.note, { ...base }));
        break;
      case 'authority':
        out.push(para(b.text, L.body.authority, { ...base, keepNext: true }));
        break;
      case 'basis':
        out.push(normal(b.text, {}, { italic: italicBasis }));
        break;
      case 'enact':
        out.push(
          para(b.text, L.body.enact, { ...base, spaceBeforePt: 6, keepNext: true }),
        );
        break;
      case 'chapter':
      case 'chapter-title':
        out.push(para(b.text, L.body.chapter, { ...base, keepNext: true }));
        break;
      case 'article': {
        const title = b.text.replace(/^Điều\s+\d+\s*[.:]?\s*/iu, '');
        const short =
          title.length > 0 &&
          wordCount(title) <= L.body.article.shortTitleMaxWords &&
          !/[.;,]$/.test(title);
        out.push(normal(b.text, short ? { keepNext: true } : {}, { bold: short }));
        break;
      }
      case 'heading': {
        const st = L.body.headings[String(b.level ?? 1)] ?? {};
        out.push(normal(b.text, { keepNext: true }, st));
        break;
      }
      default:
        out.push(normal(b.text));
    }
  }
  return out;
}

function closingBlock(m: Model, L: Layout): BlockNode[] {
  const C = L.closing;
  const items = punctuateList(m.recipients.length ? m.recipients : ['Lưu: VT']);

  const leftCell: BlockNode[] = [
    para('Nơi nhận:', C.recipientsTitle, { spaceBeforePt: C.spaceBeforePt, lineSpacing: 1 }),
    ...items.map((t) => para(t, C.recipientItem, { lineSpacing: 1 })),
  ];

  const signer = C.signer;
  const rightCell: BlockNode[] = [];
  m.signerTitle.forEach((t, i) => {
    rightCell.push(
      para(t, { ...signer, caps: true }, {
        lineSpacing: 1,
        spaceBeforePt: i === 0 ? C.spaceBeforePt : 0,
        keepNext: true,
      }),
    );
  });
  const nameNode = para(m.signerName ?? '', signer, {
    lineSpacing: 1,
    spaceBeforePt: C.signatureSpacePt,
  });
  if (m.signerTitle.length === 0) nameNode.spaceBeforePt = C.signatureSpacePt + C.spaceBeforePt;
  rightCell.push(nameNode);

  return [
    {
      type: 'table',
      cantSplit: true,
      rows: [
        [
          { children: leftCell, widthMm: C.leftColumnMm },
          { children: rightCell, widthMm: C.rightColumnMm },
        ],
      ],
    },
  ];
}

export function buildIR(m: Model, L: Layout = ND30_LAYOUT): DocumentIR {
  const P = L.page;
  const contentMm = P.widthMm - P.margins.leftMm - P.margins.rightMm;
  return {
    fontFamily: L.font,
    defaultSizePt: L.defaults.bodySizePt,
    page: {
      widthMm: P.widthMm,
      heightMm: P.heightMm,
      marginTopMm: P.margins.topMm,
      marginBottomMm: P.margins.bottomMm,
      marginLeftMm: P.margins.leftMm,
      marginRightMm: P.margins.rightMm,
      headerDistanceMm: P.headerDistanceMm,
      footerDistanceMm: P.footerDistanceMm,
      pageNumber: P.pageNumber ? { ...P.pageNumber } : null,
    },
    children: [
      ...headerBlock(m, L),
      ...titleBlock(m, L, contentMm),
      ...addresseeBlock(m, L),
      ...bodyBlocks(m, L),
      ...closingBlock(m, L),
    ],
  };
}
