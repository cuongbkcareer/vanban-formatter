/** Kiểu dữ liệu của profile văn bản hành chính (NĐ 30/2020/NĐ-CP). */

/** Các loại văn bản hành chính theo NĐ 30 (không gồm công văn: công văn không có tên loại). */
export const DOC_TYPES = [
  'Nghị quyết',
  'Quyết định',
  'Chỉ thị',
  'Quy chế',
  'Quy định',
  'Thông cáo',
  'Thông báo',
  'Hướng dẫn',
  'Chương trình',
  'Kế hoạch',
  'Phương án',
  'Đề án',
  'Dự án',
  'Báo cáo',
  'Biên bản',
  'Tờ trình',
  'Hợp đồng',
  'Công điện',
  'Bản ghi nhớ',
  'Bản cam kết',
  'Bản thỏa thuận',
  'Giấy ủy quyền',
  'Giấy mời',
  'Giấy giới thiệu',
  'Giấy nghỉ phép',
  'Phiếu gửi',
  'Phiếu chuyển',
  'Thư công',
] as const;

export const CONG_VAN = 'Công văn';

/** Nhãn gán cho từng dòng của văn bản thô. */
export type Label =
  | 'parent-org'
  | 'issuer'
  | 'national-title'
  | 'motto'
  | 'decor'
  | 'number'
  | 'place-date'
  | 'doc-type'
  | 'subject'
  | 'subject-cv'
  | 'addressee'
  | 'authority'
  | 'basis'
  | 'enact'
  | 'chapter'
  | 'chapter-title'
  | 'article'
  | 'heading1'
  | 'heading2'
  | 'heading3'
  | 'heading4'
  | 'paragraph'
  | 'recipients-title'
  | 'recipient'
  | 'signer-title'
  | 'signer-name'
  | 'signature-note';

export const LABEL_NAMES: Record<Label, string> = {
  'parent-org': 'Cơ quan chủ quản',
  issuer: 'Cơ quan ban hành',
  'national-title': 'Quốc hiệu (tự sinh lại)',
  motto: 'Tiêu ngữ (tự sinh lại)',
  decor: 'Dòng trang trí (bỏ)',
  number: 'Số, ký hiệu',
  'place-date': 'Địa danh, ngày tháng',
  'doc-type': 'Tên loại văn bản',
  subject: 'Trích yếu',
  'subject-cv': 'Trích yếu công văn (V/v)',
  addressee: 'Kính gửi',
  authority: 'Thẩm quyền ban hành',
  basis: 'Căn cứ',
  enact: 'Từ "QUYẾT ĐỊNH:"',
  chapter: 'Chương',
  'chapter-title': 'Tên chương',
  article: 'Điều',
  heading1: 'Tiêu đề cấp 1 (I.)',
  heading2: 'Tiêu đề cấp 2 (1.)',
  heading3: 'Tiêu đề cấp 3 (a))',
  heading4: 'Tiêu đề cấp 4',
  paragraph: 'Đoạn nội dung',
  'recipients-title': 'Chữ "Nơi nhận"',
  recipient: 'Nơi nhận (mục)',
  'signer-title': 'Chức vụ người ký',
  'signer-name': 'Họ tên người ký',
  'signature-note': 'Ghi chú chữ ký (bỏ)',
};

export interface LabeledLine {
  text: string;
  label: Label;
  /** 0..1 — dưới 0.7 nên được người dùng kiểm tra lại. */
  confidence: number;
  note?: string;
}

/** Thông tin người dùng nhập để bổ sung/ghi đè những gì khó đoán từ văn bản thô. */
export interface Options {
  /** Cơ quan chủ quản; nhiều dòng ngăn bằng \n. */
  parentOrg?: string;
  issuer?: string;
  place?: string;
  /** Tên loại: một trong DOC_TYPES, hoặc 'Công văn'. */
  docType?: string;
  /** YYYY-MM-DD */
  dateIso?: string;
  /** Số, ký hiệu, ví dụ 12/QĐ-UBND */
  number?: string;
  /** Chức vụ người ký; nhiều dòng ngăn bằng \n (ví dụ "TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH"). */
  signerTitle?: string;
  signerName?: string;
}

export type BodyKind =
  | 'paragraph'
  | 'authority'
  | 'basis'
  | 'enact'
  | 'chapter'
  | 'chapter-title'
  | 'article'
  | 'heading';

export interface BodyBlock {
  kind: BodyKind;
  /** Chỉ dùng cho kind = 'heading' (1..4). */
  level?: 1 | 2 | 3 | 4;
  text: string;
}

export interface PlaceDate {
  place?: string;
  day?: number;
  month?: number;
  year?: number;
}

/** Mô hình cấu trúc của văn bản — đầu ra của bước assemble, đầu vào của build-ir. */
export interface Model {
  /** undefined = công văn */
  docType?: string;
  parentOrg: string[];
  issuer?: string;
  number?: string;
  placeDate: PlaceDate;
  subject?: string;
  /** Trích yếu công văn (đã bỏ tiền tố V/v). */
  vv?: string;
  addressees: string[];
  body: BodyBlock[];
  recipients: string[];
  signerTitle: string[];
  signerName?: string;
  warnings: string[];
}
