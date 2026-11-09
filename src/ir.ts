/**
 * IR (Intermediate Representation) — lớp trung gian dùng chung cho mọi profile.
 *
 * Profile (NĐ 30, thư tiếng Anh, ...) chỉ cần sinh ra IR; renderer .docx là một.
 * Mọi kích thước dùng mm và pt để dễ đối chiếu với văn bản quy định.
 */

export interface Run {
  text: string;
  bold?: boolean;
  italic?: boolean;
  /** Cỡ chữ (pt). Mặc định lấy từ Paragraph.sizePt. */
  sizePt?: number;
}

export type Align = 'left' | 'center' | 'right' | 'justify';

export interface ParagraphNode {
  type: 'paragraph';
  runs: Run[];
  align?: Align;
  sizePt?: number;
  firstLineMm?: number;
  leftMm?: number;
  rightMm?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  /** Giãn dòng dạng bội số (1 = đơn). */
  lineSpacing?: number;
  keepNext?: boolean;
  keepLines?: boolean;
}

/** Đường kẻ ngang (nét liền) có độ dài xác định, căn giữa trong khung chứa nó. */
export interface RuleNode {
  type: 'rule';
  widthMm: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
}

export interface TableCellNode {
  children: BlockNode[];
  widthMm: number;
}

export interface TableNode {
  type: 'table';
  /** Bảng không viền, dùng để dàn trang (phần đầu, phần nơi nhận / chữ ký). */
  rows: TableCellNode[][];
  cantSplit?: boolean;
}

export type BlockNode = ParagraphNode | RuleNode | TableNode;

export interface PageSetup {
  widthMm: number;
  heightMm: number;
  marginTopMm: number;
  marginBottomMm: number;
  marginLeftMm: number;
  marginRightMm: number;
  headerDistanceMm: number;
  footerDistanceMm: number;
  /** Số trang: căn giữa, ở lề trên; null nếu không đánh số. */
  pageNumber: null | {
    sizePt: number;
    /** Không hiển thị số trang trên trang đầu. */
    hideOnFirstPage: boolean;
  };
}

export interface DocumentIR {
  fontFamily: string;
  defaultSizePt: number;
  page: PageSetup;
  children: BlockNode[];
}

/** Chiều rộng vùng nội dung (mm). */
export function contentWidthMm(page: PageSetup): number {
  return page.widthMm - page.marginLeftMm - page.marginRightMm;
}
