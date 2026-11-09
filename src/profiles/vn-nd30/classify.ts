import {
  capitalizeFirst,
  isAllCaps,
  joinHardWraps,
  normalizeText,
  splitHeaderColumns,
  toLowerVi,
  wordCount,
} from '../../text';
import { DOC_TYPES, type Label, type LabeledLine, type PlaceDate } from './types';

/* ------------------------------------------------------------------ */
/* Biểu thức chính quy                                                 */
/* ------------------------------------------------------------------ */

const typeAlt = DOC_TYPES.map((t) => t.replace(/ /g, '\\s+')).join('|');
const RE_TYPE_EXACT = new RegExp(`^(${typeAlt})\\s*([:：])?$`, 'iu');
const RE_TYPE_PREFIX = new RegExp(`^(${typeAlt})\\s+(.+)$`, 'iu');

const RE_DECOR = /^[-_*=~.\s]{3,}$/u;
const RE_NATIONAL = /^CỘNG\s+H(?:ÒA|OÀ)\s+XÃ\s+HỘI\s+CHỦ\s+NGHĨA\s+VIỆT\s+NAM$/iu;
const RE_MOTTO = /^Độc\s+lập\s*[-–—]\s*Tự\s+do\s*[-–—]\s*Hạnh\s+phúc[\s\-–—_.*]*$/iu;
const RE_NUMBER = /^Số\s*[:：]\s*\S.*$|^Số\s*[:：]\s*$|^Số\s+\d+\s*\/\S*$/iu;
const RE_VV = /^(V\/v|Về\s+việc)\s*[:.]?\s*(.*)$/iu;
const RE_KINH_GUI = /^Kính\s+gửi\s*[:：]?\s*(.*)$/iu;
const RE_BASIS = /^Căn\s+cứ\s+\S/iu;
const RE_CHAPTER = /^Chương\s+([IVXLC]+|\d+)\b\s*[.:]?\s*(.*)$/iu;
const RE_ARTICLE = /^Điều\s+\d+\s*([.:]|$)/iu;
const RE_BULLET = /^[-–—+•*]\s*/u;
const RE_RECIPIENTS = /^Nơi\s+nhận\s*[:：]?\s*$/iu;
const RE_RECIPIENT_ITEM = /^(Như\s+trên|Lưu\s*[:：]?)/iu;
const RE_SIGN_NOTE =
  /^\(?\s*(đã\s+ký|ký\s+tên|chữ\s+ký|ký\s*,?\s*(ghi\s+rõ\s+họ\s+tên|đóng\s+dấu)|ký\s+và\s+đóng\s+dấu)[^)]*\)?$/iu;

const RE_H1 = /^([IVXLC]+)\s*[.)]\s+\S/u;
const RE_H_DECIMAL = /^(\d+(?:\.\d+)+)\.?\s+\S/u;
const RE_H2 = /^(\d+)\s*[.)]\s+\S/u;
const RE_H3 = /^([a-zđ])\s*\)\s+\S/iu;
const RE_HEADING_START = /^([IVXLC]+\s*[.)]|\d+\s*[.)]|[a-zđ]\s*\))\s/iu;

const RE_POWER = /^(?:TM|KT|TL|TUQ|Q)\.\s*\S/u;
const RE_ROLE =
  /^(?:(?:TM|KT|TL|TUQ|Q)\.\s*)?(?:Phó\s+)?(?:Chủ\s+tịch|Giám\s+đốc|Tổng\s+giám\s+đốc|Trưởng\s+\p{L}+|Bí\s+thư|Chủ\s+nhiệm|Hiệu\s+trưởng|Chánh\s+\p{L}+|Thủ\s+trưởng|Bộ\s+trưởng|Thứ\s+trưởng|Cục\s+trưởng|Viện\s+trưởng|Ủy\s+viên|Thư\s+ký|Quyền\s+\p{L}+|Đại\s+diện)/iu;

const RE_ORG_GENERIC = /^(ỦY\s+BAN\s+NHÂN\s+DÂN|HỘI\s+ĐỒNG\s+NHÂN\s+DÂN|UBND|HĐND)$/iu;
const RE_GEO_START = /^(TỈNH|THÀNH\s+PHỐ|TP\.?|HUYỆN|QUẬN|THỊ\s+XÃ|THỊ\s+TRẤN|XÃ|PHƯỜNG)\s+\S/iu;
const RE_ORG_WORD =
  /^(UBND|Ủy\s+ban|Sở|Bộ|Công\s+ty|Trường|Văn\s+phòng|Phòng|Ban|Hội\s+đồng|Chi\s+cục|Cục|Tổng\s+cục|Viện|Trung\s+tâm|Đảng\s+ủy|Chi\s+bộ|Đoàn|Hội|Liên\s+đoàn|Quốc\s+hội|Chính\s+phủ|Tòa\s+án|Công\s+an|Bệnh\s+viện|Tập\s+đoàn|Tổng\s+công\s+ty|Thanh\s+tra|Kho\s+bạc|Thành\s+ủy|Tỉnh\s+ủy|Huyện\s+ủy|Học\s+viện|Đại\s+học)(?!\p{L})/iu;

/* ------------------------------------------------------------------ */
/* Địa danh, ngày tháng                                                */
/* ------------------------------------------------------------------ */

const RE_DATE_LONG =
  /^(?:(.+?)\s*,\s*)?(?:ngày\s+)?([\d.…_]+)\s*tháng\s*([\d.…_]+)\s*năm\s*([\d.…_]+)\s*\.?$/iu;
const RE_DATE_NUM = /^(?:(.+?)\s*,\s*)?(?:ngày\s+)?(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})\.?$/iu;

function num(s: string | undefined): number | undefined {
  return s && /^\d+$/.test(s) ? Number(s) : undefined;
}

/** Nhận diện dòng "Địa danh, ngày … tháng … năm …". Trả về null nếu không khớp. */
export function parsePlaceDate(line: string): PlaceDate | null {
  const m = RE_DATE_LONG.exec(line) ?? RE_DATE_NUM.exec(line);
  if (!m) return null;
  // "ngày 5 tháng 3" phải có cụm "ngày" hoặc "tháng" rõ ràng; bỏ các chuỗi số thuần.
  if (!/tháng|[/.\-]\d/iu.test(line)) return null;
  const pd: PlaceDate = {};
  if (m[1]) pd.place = m[1].trim();
  const day = num(m[2]);
  const month = num(m[3]);
  const year = num(m[4]);
  if (day !== undefined) pd.day = day;
  if (month !== undefined) pd.month = month;
  if (year !== undefined) pd.year = year;
  return pd;
}

/* ------------------------------------------------------------------ */
/* Phần cuối: nơi nhận và chữ ký                                       */
/* ------------------------------------------------------------------ */

function isTitleLike(line: string): boolean {
  if (wordCount(line) > 9 || RE_HEADING_START.test(line)) return false;
  return RE_POWER.test(line) || RE_ROLE.test(line);
}

function isNameLike(line: string): boolean {
  const wc = wordCount(line);
  if (wc < 2 || wc > 6) return false;
  if (/\d/.test(line) || RE_BULLET.test(line) || /[;:,!?]$/.test(line)) return false;
  return line.split(/\s+/).every((w) => /^\p{Lu}/u.test(w));
}

interface SignerHit {
  titles: number[];
  name?: number;
  notes: number[];
}

function scanSigner(lines: string[], end: number, minStart: number): SignerHit | null {
  const notes: number[] = [];
  let i = end - 1;
  const skipNotes = (): void => {
    while (i >= minStart && RE_SIGN_NOTE.test(lines[i] ?? '')) {
      notes.push(i);
      i--;
    }
  };
  skipNotes();
  if (i < minStart) return null;
  let name: number | undefined;
  const last = lines[i] ?? '';
  if (!isTitleLike(last) && isNameLike(last)) {
    name = i;
    i--;
    skipNotes();
  }
  const titles: number[] = [];
  while (i >= minStart && titles.length < 4 && isTitleLike(lines[i] ?? '')) {
    titles.unshift(i);
    i--;
  }
  if (titles.length === 0) return null;
  return name === undefined ? { titles, notes } : { titles, name, notes };
}

function detectTail(lines: string[]): Map<number, LabeledLine> {
  const tail = new Map<number, LabeledLine>();
  const n = lines.length;
  let rIdx = -1;
  for (let i = n - 1; i >= 0; i--) {
    if (RE_RECIPIENTS.test(lines[i] ?? '')) {
      rIdx = i;
      break;
    }
  }

  let recipientsEnd = 0;
  if (rIdx >= 0) {
    tail.set(rIdx, { text: lines[rIdx] ?? '', label: 'recipients-title', confidence: 0.98 });
    let j = rIdx + 1;
    while (j < n) {
      const l = lines[j] ?? '';
      if (RE_BULLET.test(l) || RE_RECIPIENT_ITEM.test(l)) {
        tail.set(j, { text: l, label: 'recipient', confidence: 0.95 });
        j++;
      } else break;
    }
    recipientsEnd = j;
  }

  // Chữ ký có thể nằm sau danh sách nơi nhận, hoặc ngay trước chữ "Nơi nhận".
  let hit: SignerHit | null = null;
  if (rIdx >= 0) {
    hit = scanSigner(lines, n, recipientsEnd);
    if (!hit) hit = scanSigner(lines, rIdx, 0);
  } else {
    hit = scanSigner(lines, n, 0);
  }
  if (hit) {
    for (const t of hit.titles)
      tail.set(t, { text: lines[t] ?? '', label: 'signer-title', confidence: 0.85 });
    if (hit.name !== undefined)
      tail.set(hit.name, { text: lines[hit.name] ?? '', label: 'signer-name', confidence: 0.8 });
    for (const k of hit.notes)
      tail.set(k, { text: lines[k] ?? '', label: 'signature-note', confidence: 0.9 });
  }
  return tail;
}

/* ------------------------------------------------------------------ */
/* Phân loại chính                                                     */
/* ------------------------------------------------------------------ */

function isShortNoPeriod(line: string, maxWords: number): boolean {
  return wordCount(line) <= maxWords && !/[.;,:]$/.test(line);
}

function looksLikeOrg(line: string, position: number): { ok: boolean; confidence: number } {
  if (wordCount(line) > 12 || /[.;:]$/.test(line)) return { ok: false, confidence: 0 };
  if (isAllCaps(line)) return { ok: true, confidence: RE_ORG_WORD.test(line) ? 0.9 : 0.6 };
  if (position < 6 && wordCount(line) <= 10 && RE_ORG_WORD.test(line))
    return { ok: true, confidence: 0.65 };
  return { ok: false, confidence: 0 };
}

export function classify(raw: string): LabeledLine[] {
  const lines = joinHardWraps(splitHeaderColumns(normalizeText(raw)));
  const tail = detectTail(lines);
  const hasArticles = lines.some((l, i) => !tail.has(i) && RE_ARTICLE.test(l));

  const out: LabeledLine[] = [];
  const orgIdx: number[] = [];
  let zone: 'header' | 'body' = 'header';
  let addr: 'none' | 'expect-first' | 'bullets' = 'none';
  let chapterTitleNext = false;
  const seen = { docType: false, subject: false, number: false, date: false, basis: false };

  const push = (label: Label, text: string, confidence = 0.95, note?: string): void => {
    out.push(note ? { label, text, confidence, note } : { label, text, confidence });
  };

  for (let i = 0; i < lines.length; i++) {
    const preset = tail.get(i);
    if (preset) {
      out.push(preset);
      continue;
    }
    const line = lines[i] ?? '';

    if (RE_DECOR.test(line)) {
      push('decor', line, 0.99);
      continue;
    }
    if (RE_NATIONAL.test(line)) {
      push('national-title', line, 0.99);
      continue;
    }
    if (RE_MOTTO.test(line)) {
      push('motto', line, 0.99);
      continue;
    }

    // Danh sách "Kính gửi" nhiều dòng
    if (addr !== 'none') {
      const isBullet = RE_BULLET.test(line);
      if (isBullet) {
        push('addressee', line.replace(RE_BULLET, ''), 0.9);
        addr = 'bullets';
        continue;
      }
      if (addr === 'expect-first' && wordCount(line) <= 25 && !RE_BASIS.test(line)) {
        push('addressee', line, 0.8);
        addr = 'none';
        continue;
      }
      addr = 'none';
    }

    if (chapterTitleNext) {
      chapterTitleNext = false;
      if (isShortNoPeriod(line, 15) && !RE_ARTICLE.test(line) && !RE_CHAPTER.test(line)) {
        push('chapter-title', line, 0.85);
        continue;
      }
    }

    if (!seen.number && zone === 'header' && RE_NUMBER.test(line)) {
      push('number', line);
      seen.number = true;
      continue;
    }
    if (!seen.date) {
      const pd = parsePlaceDate(line);
      if (pd) {
        push('place-date', line, 0.95);
        seen.date = true;
        continue;
      }
    }

    // "QUYẾT ĐỊNH:" sau các dòng Căn cứ
    const exact = RE_TYPE_EXACT.exec(line);
    if (exact && seen.basis && (seen.docType || exact[2])) {
      push('enact', line, 0.9);
      continue;
    }

    if (RE_BASIS.test(line)) {
      push('basis', line, 0.9);
      seen.basis = true;
      zone = 'body';
      continue;
    }

    if (zone === 'header') {
      const vv = RE_VV.exec(line);
      if (vv && !seen.subject) {
        const typedDoc = seen.docType && /^Về\s+việc/iu.test(line);
        push(typedDoc ? 'subject' : 'subject-cv', line, 0.95);
        seen.subject = true;
        continue;
      }
      if (exact && !seen.docType) {
        push('doc-type', line, 0.95);
        seen.docType = true;
        continue;
      }
      const prefix = RE_TYPE_PREFIX.exec(line);
      if (prefix && !seen.docType && isAllCaps(line)) {
        const typeText = (prefix[1] ?? '').trim();
        const rest = (prefix[2] ?? '').trim();
        push('doc-type', typeText, 0.8, 'Tách từ dòng "' + line + '"');
        push('subject', capitalizeFirst(toLowerVi(rest)), 0.5, 'Trích yếu tách từ tên loại (đã hạ chữ thường), cần kiểm tra viết hoa tên riêng');
        seen.docType = true;
        seen.subject = true;
        continue;
      }
      const kg = RE_KINH_GUI.exec(line);
      if (kg) {
        const rest = (kg[1] ?? '').trim();
        if (rest) {
          push('addressee', rest.replace(RE_BULLET, ''), 0.95);
          addr = 'bullets';
        } else {
          addr = 'expect-first';
        }
        continue;
      }
      if (!seen.docType && !seen.number && !seen.subject && orgIdx.length < 3) {
        const org = looksLikeOrg(line, out.length);
        if (org.ok) {
          orgIdx.push(out.length);
          push('issuer', line, org.confidence);
          continue;
        }
      }
      if (seen.docType && seen.subject && !seen.basis && isAllCaps(line) && wordCount(line) <= 12 && !/[.;:]$/.test(line)) {
        push('authority', line, 0.8);
        continue;
      }
      if (
        seen.docType &&
        !seen.subject &&
        isShortNoPeriod(line, 30) &&
        !RE_HEADING_START.test(line) &&
        !RE_ARTICLE.test(line)
      ) {
        push('subject', line, 0.8);
        seen.subject = true;
        continue;
      }
      zone = 'body';
    }

    /* ---------------- Thân văn bản ---------------- */
    const chap = RE_CHAPTER.exec(line);
    if (chap) {
      push('chapter', `Chương ${chap[1] ?? ''}`.trim(), 0.95);
      const rest = (chap[2] ?? '').trim();
      if (rest) push('chapter-title', rest, 0.85);
      else chapterTitleNext = true;
      continue;
    }
    if (RE_ARTICLE.test(line)) {
      push('article', line, 0.95);
      continue;
    }
    if (RE_H1.test(line) && wordCount(line) <= 20) {
      push('heading1', line, 0.85);
      continue;
    }
    if (!hasArticles && RE_H_DECIMAL.test(line) && isShortNoPeriod(line, 20)) {
      push('heading3', line, 0.7);
      continue;
    }
    if (RE_H2.test(line)) {
      if (!hasArticles && isShortNoPeriod(line, 20)) push('heading2', line, 0.75);
      else push('paragraph', line, 0.85);
      continue;
    }
    if (RE_H3.test(line)) {
      if (!hasArticles && isShortNoPeriod(line, 20)) push('heading3', line, 0.7);
      else push('paragraph', line, 0.85);
      continue;
    }
    push('paragraph', line, 0.85);
  }

  // "ỦY BAN NHÂN DÂN" xuống dòng "THÀNH PHỐ MẪU" là MỘT tên cơ quan, không phải chủ quản + ban hành.
  for (let k = orgIdx.length - 2; k >= 0; k--) {
    const a = out[orgIdx[k] ?? -1];
    const b = out[orgIdx[k + 1] ?? -1];
    if (a && b && orgIdx[k + 1] === (orgIdx[k] ?? -9) + 1 && RE_ORG_GENERIC.test(a.text) && RE_GEO_START.test(b.text)) {
      b.text = `${a.text} ${b.text}`;
      b.confidence = a.confidence;
      out.splice(orgIdx[k] as number, 1);
      for (let j = k + 1; j < orgIdx.length; j++) orgIdx[j] = (orgIdx[j] as number) - 1;
      orgIdx.splice(k, 1);
    }
  }

  // Dòng tên cơ quan: dòng cuối là cơ quan ban hành, các dòng trước là cơ quan chủ quản.
  orgIdx.forEach((idx, k) => {
    const item = out[idx];
    if (item && k < orgIdx.length - 1) item.label = 'parent-org';
  });

  return out;
}
