/** Tiện ích xử lý văn bản thô — dùng chung cho các profile. */

const VI = 'vi';

export function toUpperVi(s: string): string {
  return s.toLocaleUpperCase(VI);
}

export function toLowerVi(s: string): string {
  return s.toLocaleLowerCase(VI);
}

export function capitalizeFirst(s: string): string {
  if (!s) return s;
  return toUpperVi(s.charAt(0)) + s.slice(1);
}

/** Dòng có chữ cái và toàn bộ chữ cái đều là chữ hoa. */
export function isAllCaps(s: string): boolean {
  const letters = s.replace(/[^\p{L}]/gu, '');
  return letters.length > 0 && letters === toUpperVi(letters);
}

export function wordCount(s: string): number {
  const t = s.trim();
  return t ? t.split(/\s+/).length : 0;
}

export function startsWithLower(s: string): boolean {
  const c = s.charAt(0);
  return /\p{Ll}/u.test(c);
}

/** Ước lượng bề rộng chuỗi (mm) với font Times New Roman, dùng để căn đường kẻ. */
export function estimateWidthMm(text: string, sizePt: number, bold = false): number {
  let em = 0;
  for (const ch of text.normalize('NFC')) {
    if (ch === ' ') em += 0.25;
    else if (/\p{Lu}/u.test(ch)) em += 0.7;
    else if (/\p{Ll}/u.test(ch)) em += 0.45;
    else if (/\d/.test(ch)) em += 0.5;
    else em += 0.33;
  }
  if (bold) em *= 1.05;
  return (em * sizePt * 25.4) / 72;
}

/**
 * Chuẩn hoá văn bản thô:
 * - NFC, bỏ ký tự điều khiển/NBSP, cắt khoảng trắng thừa
 * - bỏ dòng trống (thể thức hành chính không dùng dòng trống để giãn đoạn)
 * - nối các dòng bị ngắt cứng giữa câu
 * Trả về danh sách dòng (mỗi dòng = một đoạn).
 */
export function normalizeText(raw: string): string[] {
  const text = raw
    .normalize('NFC')
    .replace(/\r\n?/g, '\n')
    .replace(/[   ]/g, ' ')
    .replace(/[​-‍﻿]/g, '')
    .replace(/\f/g, '\n');

  const out: string[] = [];
  const rawLines = text.split('\n');
  rawLines.forEach((line, idx) => {
    // Giữ nguyên tab ở 12 dòng đầu để tách cột của phần đầu văn bản.
    const l = (idx < 12 ? line : line.replace(/\t/g, ' ')).replace(/[ ]+$/g, '').replace(/^[ ]+/g, '');
    if (l.trim() === '') return;
    out.push(l);
  });
  return out;
}

/** Tách các "cột" của phần đầu (dán từ bảng Word thường có tab / nhiều dấu cách). */
export function splitHeaderColumns(lines: string[], maxLines = 12): string[] {
  const out: string[] = [];
  lines.forEach((line, idx) => {
    if (idx < maxLines) {
      const parts = line
        .split(/\t+| {3,}/)
        .map((p) => p.trim())
        .filter(Boolean);
      out.push(...parts);
    } else {
      out.push(line.replace(/\t+/g, ' ').replace(/ {2,}/g, ' '));
    }
  });
  return out.map((l) => l.replace(/ {2,}/g, ' '));
}

const TERMINAL_PUNCT = /[.;:!?…]$/;

/**
 * Nối dòng bị ngắt cứng: dòng sau bắt đầu bằng chữ thường và dòng trước đủ dài,
 * không kết thúc bằng dấu câu.
 */
export function joinHardWraps(lines: string[]): string[] {
  const out: string[] = [];
  for (const line of lines) {
    const prev = out[out.length - 1];
    if (
      prev !== undefined &&
      startsWithLower(line) &&
      prev.length >= 40 &&
      !TERMINAL_PUNCT.test(prev)
    ) {
      out[out.length - 1] = `${prev} ${line}`;
    } else {
      out.push(line);
    }
  }
  return out;
}
