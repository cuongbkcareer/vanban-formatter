import {
  AlignmentType,
  BorderStyle,
  Document,
  Header,
  LineRuleType,
  Packer,
  PageNumber,
  ShadingType,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
  type ParagraphChild,
} from 'docx';
import type {
  Align,
  BlockNode,
  BorderSpec,
  DocumentIR,
  ParagraphNode,
  RuleNode,
  TableNode,
} from './ir';
import { contentWidthMm } from './ir';

/** 1 mm = 1440/25.4 twips */
const mmToTwip = (mm: number): number => Math.round((mm * 1440) / 25.4);
const ptToTwip = (pt: number): number => Math.round(pt * 20);
const ptToHalf = (pt: number): number => Math.round(pt * 2);

const ALIGN: Record<Align, (typeof AlignmentType)[keyof typeof AlignmentType]> = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.BOTH,
};

interface Ctx {
  font: string;
  defaultSizePt: number;
}

function renderParagraph(node: ParagraphNode, ctx: Ctx): Paragraph {
  const size = node.sizePt ?? ctx.defaultSizePt;
  const children: ParagraphChild[] = node.runs.map(
    (r) =>
      new TextRun({
        text: r.text,
        bold: r.bold,
        italics: r.italic,
        size: ptToHalf(r.sizePt ?? size),
        font: ctx.font,
      }),
  );
  return new Paragraph({
    children,
    alignment: ALIGN[node.align ?? 'left'],
    keepNext: node.keepNext,
    keepLines: node.keepLines,
    indent: {
      firstLine: node.firstLineMm ? mmToTwip(node.firstLineMm) : undefined,
      left: node.leftMm ? mmToTwip(node.leftMm) : undefined,
      right: node.rightMm ? mmToTwip(node.rightMm) : undefined,
    },
    spacing: {
      before: ptToTwip(node.spaceBeforePt ?? 0),
      after: ptToTwip(node.spaceAfterPt ?? 0),
      line: Math.round((node.lineSpacing ?? 1) * 240),
      lineRule: LineRuleType.AUTO,
    },
  });
}

/**
 * Đường kẻ = đoạn rỗng cỡ nhỏ có viền dưới; độ dài điều khiển bằng thụt lề hai bên
 * để căn giữa trong khung chứa (trang hoặc ô bảng).
 */
function renderRule(node: RuleNode, containerMm: number, ctx: Ctx): Paragraph {
  const inset = Math.max(0, (containerMm - node.widthMm) / 2);
  return new Paragraph({
    children: [new TextRun({ text: '', size: 2, font: ctx.font })],
    run: { size: 2 },
    alignment: AlignmentType.CENTER,
    indent: { left: mmToTwip(inset), right: mmToTwip(inset) },
    spacing: {
      before: ptToTwip(node.spaceBeforePt ?? 0),
      after: ptToTwip(node.spaceAfterPt ?? 0),
      line: 20,
      lineRule: LineRuleType.EXACT,
    },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, space: 0, color: '000000' },
    },
  });
}

const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' } as const;
const NO_BORDERS = {
  top: NO_BORDER,
  bottom: NO_BORDER,
  left: NO_BORDER,
  right: NO_BORDER,
};

const BORDER_STYLE = {
  single: BorderStyle.SINGLE,
  double: BorderStyle.DOUBLE,
  dashed: BorderStyle.DASHED,
  dotted: BorderStyle.DOTTED,
} as const;

function border(b: BorderSpec | null | undefined) {
  if (!b) return NO_BORDER;
  return { style: BORDER_STYLE[b.style ?? 'single'], size: Math.max(2, Math.round(b.sizePt * 8)), color: b.color };
}

const V_ALIGN = {
  top: VerticalAlign.TOP,
  center: VerticalAlign.CENTER,
  bottom: VerticalAlign.BOTTOM,
} as const;

function renderTable(node: TableNode, ctx: Ctx): Table {
  const widths = (node.rows[0] ?? []).map((c) => mmToTwip(c.widthMm));
  const total = widths.reduce((a, b) => a + b, 0);
  const fm = node.format;
  const b = fm?.borders;
  const mar = fm?.cellMarginMm;
  const cellMargins = mar
    ? { top: mmToTwip(mar[0]), right: mmToTwip(mar[1]), bottom: mmToTwip(mar[2]), left: mmToTwip(mar[3]) }
    : { top: 0, bottom: 0, left: mmToTwip(1.5), right: mmToTwip(1.5) };
  const cellBorders = {
    top: NO_BORDER,
    bottom: NO_BORDER,
    left: NO_BORDER,
    right: NO_BORDER,
  };
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    // Lề trong ô đặt ở mức bảng (w:tblCellMar) để Word hiển thị và đọc lại đúng như một thuộc tính của bảng.
    margins: { marginUnitType: WidthType.DXA, ...cellMargins },
    ...(fm?.align ? { alignment: ALIGN[fm.align] } : {}),
    borders: {
      top: border(b?.top),
      bottom: border(b?.bottom),
      left: border(b?.left),
      right: border(b?.right),
      insideHorizontal: border(b?.insideH),
      insideVertical: border(b?.insideV),
    },
    rows: node.rows.map(
      (cells, rowIdx) =>
        new TableRow({
          cantSplit: node.cantSplit,
          ...(rowIdx === 0 && fm?.headerRepeat ? { tableHeader: true } : {}),
          children: cells.map(
            (cell) =>
              new TableCell({
                width: { size: mmToTwip(cell.widthMm), type: WidthType.DXA },
                // Viền của bảng được áp bằng tblBorders; ô không ghi đè để Word dùng viền bảng.
                ...(fm?.borders ? {} : { borders: cellBorders }),
                verticalAlign: V_ALIGN[fm?.vAlign ?? 'top'],
                ...(rowIdx === 0 && fm?.headerShading
                  ? { shading: { type: ShadingType.CLEAR, fill: fm.headerShading, color: 'auto' } }
                  : {}),
                children: renderBlocks(cell.children, cell.widthMm, ctx),
              }),
          ),
        }),
    ),
  });
}

function renderBlocks(
  blocks: BlockNode[],
  containerMm: number,
  ctx: Ctx,
): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];
  for (const b of blocks) {
    if (b.type === 'paragraph') out.push(renderParagraph(b, ctx));
    else if (b.type === 'rule') out.push(renderRule(b, containerMm, ctx));
    else out.push(renderTable(b, ctx));
  }
  // Ô bảng bắt buộc có ít nhất một đoạn.
  if (out.length === 0) out.push(new Paragraph({}));
  return out;
}

function pageNumberHeader(ir: DocumentIR): Header {
  const pn = ir.page.pageNumber;
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: pn
          ? [
              new TextRun({
                children: [PageNumber.CURRENT],
                size: ptToHalf(pn.sizePt),
                font: ir.fontFamily,
              }),
            ]
          : [],
      }),
    ],
  });
}

function buildDocument(ir: DocumentIR): Document {
  const ctx: Ctx = { font: ir.fontFamily, defaultSizePt: ir.defaultSizePt };
  const p = ir.page;
  const pn = p.pageNumber;
  const emptyHeader = new Header({ children: [new Paragraph({})] });

  return new Document({
    creator: 'vanban-formatter',
    styles: {
      default: {
        document: {
          run: { font: ir.fontFamily, size: ptToHalf(ir.defaultSizePt) },
        },
      },
    },
    sections: [
      {
        properties: {
          titlePage: pn?.hideOnFirstPage ?? false,
          page: {
            size: { width: mmToTwip(p.widthMm), height: mmToTwip(p.heightMm) },
            margin: {
              top: mmToTwip(p.marginTopMm),
              bottom: mmToTwip(p.marginBottomMm),
              left: mmToTwip(p.marginLeftMm),
              right: mmToTwip(p.marginRightMm),
              header: mmToTwip(p.headerDistanceMm),
              footer: mmToTwip(p.footerDistanceMm),
            },
          },
        },
        headers: pn
          ? { default: pageNumberHeader(ir), first: emptyHeader }
          : undefined,
        children: renderBlocks(ir.children, contentWidthMm(p), ctx),
      },
    ],
  });
}

export async function renderDocx(ir: DocumentIR): Promise<Uint8Array> {
  const buf = await Packer.toBuffer(buildDocument(ir));
  return new Uint8Array(buf);
}

/** Dùng cho Office.js (Word.createDocument / insertFileFromBase64 nhận chuỗi base64). */
export async function renderDocxBase64(ir: DocumentIR): Promise<string> {
  return Packer.toBase64String(buildDocument(ir));
}

/** Dùng trong trình duyệt (bản web offline). */
export async function renderDocxBlob(ir: DocumentIR): Promise<Blob> {
  return Packer.toBlob(buildDocument(ir));
}
