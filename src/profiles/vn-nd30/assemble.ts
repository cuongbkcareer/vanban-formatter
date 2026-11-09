import { toUpperVi } from '../../text';
import { parsePlaceDate } from './classify';
import {
  CONG_VAN,
  DOC_TYPES,
  type BodyBlock,
  type LabeledLine,
  type Model,
  type Options,
} from './types';

const squash = (s: string): string => s.replace(/\s+/g, ' ').trim().toLocaleLowerCase('vi');

function canonicalDocType(text: string): string {
  const t = text.replace(/[:：]\s*$/, '');
  return DOC_TYPES.find((d) => squash(d) === squash(t)) ?? t.trim();
}

function parseNumber(text: string): string {
  return text
    .replace(/^Số\s*[:：]?\s*/iu, '')
    .replace(/\s*\/\s*/g, '/')
    .replace(/\s*-\s*/g, '-')
    .trim();
}

function stripVV(text: string): string {
  return text.replace(/^(V\/v|Về\s+việc)\s*[:.]?\s*/iu, '').trim();
}

const stripBullet = (s: string): string => s.replace(/^[-–—+•*]\s*/u, '').trim();

/** Gộp các dòng đã gán nhãn (có thể đã được người dùng sửa) cùng thông tin nhập thêm thành Model. */
export function assemble(lines: LabeledLine[], opts: Options = {}): Model {
  const m: Model = {
    parentOrg: [],
    placeDate: {},
    addressees: [],
    body: [],
    recipients: [],
    signerTitle: [],
    warnings: [],
  };
  const issuers: string[] = [];

  for (const l of lines) {
    const text = l.text.trim();
    switch (l.label) {
      case 'parent-org':
        m.parentOrg.push(text);
        break;
      case 'issuer':
        issuers.push(text);
        break;
      case 'number':
        m.number = parseNumber(text);
        break;
      case 'place-date': {
        const pd = parsePlaceDate(text);
        if (pd) m.placeDate = pd;
        break;
      }
      case 'doc-type':
        m.docType = canonicalDocType(text);
        break;
      case 'subject':
        m.subject = text;
        break;
      case 'subject-cv':
        m.vv = stripVV(text);
        break;
      case 'addressee':
        if (text) m.addressees.push(stripBullet(text));
        break;
      case 'basis':
        m.body.push({ kind: 'basis', text });
        break;
      case 'enact':
        m.body.push({ kind: 'enact', text: toUpperVi(text.replace(/[:：]\s*$/, '')) + ':' });
        break;
      case 'authority':
      case 'chapter':
      case 'chapter-title':
      case 'article':
      case 'paragraph':
        m.body.push({ kind: l.label, text });
        break;
      case 'heading1':
      case 'heading2':
      case 'heading3':
      case 'heading4': {
        const level = Number(l.label.slice(-1)) as 1 | 2 | 3 | 4;
        m.body.push({ kind: 'heading', level, text } satisfies BodyBlock);
        break;
      }
      case 'recipient':
        m.recipients.push(stripBullet(text));
        break;
      case 'signer-title':
        m.signerTitle.push(text);
        break;
      case 'signer-name':
        m.signerName = text;
        break;
      default:
        // national-title, motto, decor, recipients-title, signature-note: tự sinh lại hoặc bỏ.
        break;
    }
  }

  if (issuers.length > 0) {
    m.issuer = issuers[issuers.length - 1];
    m.parentOrg.push(...issuers.slice(0, -1));
  }

  applyOptions(m, opts);
  if (m.docType === CONG_VAN) m.docType = undefined;
  if (m.docType && !m.subject && m.vv) {
    m.subject = m.vv.startsWith('V') ? m.vv : `Về việc ${m.vv}`;
    m.vv = undefined;
  }

  collectWarnings(m, lines);
  return m;
}

function applyOptions(m: Model, o: Options): void {
  if (o.parentOrg?.trim())
    m.parentOrg = o.parentOrg
      .split('\n')
      .map((s) => s.trim())
      .filter(Boolean);
  if (o.issuer?.trim()) m.issuer = o.issuer.trim();
  if (o.docType?.trim()) m.docType = canonicalDocType(o.docType);
  if (o.number?.trim()) m.number = o.number.trim();
  if (o.place?.trim()) m.placeDate.place = o.place.trim();
  if (o.dateIso) {
    const d = /^(\d{4})-(\d{2})-(\d{2})$/.exec(o.dateIso);
    if (d) {
      m.placeDate.year = Number(d[1]);
      m.placeDate.month = Number(d[2]);
      m.placeDate.day = Number(d[3]);
    }
  }
  if (o.signerTitle?.trim())
    m.signerTitle = o.signerTitle
      .split('\n')
      .map((s) => s.trim())
      .filter(Boolean);
  if (o.signerName?.trim()) m.signerName = o.signerName.trim();
}

function collectWarnings(m: Model, lines: LabeledLine[]): void {
  const w = m.warnings;
  if (!m.issuer) w.push('Chưa xác định được cơ quan ban hành. Hãy nhập ở ô "Cơ quan ban hành".');
  if (!m.number) w.push('Thiếu số, ký hiệu văn bản (đã để trống chỗ điền).');
  const pd = m.placeDate;
  if (!pd.place) w.push('Thiếu địa danh.');
  if (pd.day === undefined || pd.month === undefined || pd.year === undefined)
    w.push('Thiếu ngày, tháng, năm ban hành (đã để trống chỗ điền).');
  if (m.docType && !m.subject) w.push(`Thiếu trích yếu của ${m.docType.toLocaleLowerCase('vi')}.`);
  if (!m.docType && !m.vv) w.push('Công văn thiếu trích yếu "V/v …".');
  if (m.signerTitle.length === 0) w.push('Chưa nhận diện được chức vụ người ký.');
  if (!m.signerName) w.push('Chưa nhận diện được họ tên người ký (chừa chỗ ký).');
  if (m.recipients.length === 0)
    w.push('Thiếu nơi nhận (đã thêm "Lưu: VT." làm mẫu, hãy bổ sung đầy đủ).');
  if (m.body.length === 0) w.push('Không có nội dung văn bản.');
  const low = lines.filter((l) => l.confidence < 0.7).length;
  if (low > 0) w.push(`${low} dòng được nhận diện với độ tin cậy thấp, hãy kiểm tra ở bước xem trước.`);
  for (const l of lines) if (l.note) w.push(l.note);
}
