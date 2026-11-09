import { FIELD_NAMES, showValue, TABLE_FIELD_NAMES, type CheckResult } from '../src/learn/profile';
import type { Host } from './ui';

export interface ViewOptions {
  /** Câu tóm tắt khi đạt. */
  okText: string;
  /** Câu tóm tắt khi có lỗi. */
  badText: string;
  host: Host;
  /** Kết quả lấy từ tài liệu đang mở: cho phép bấm vào lỗi để chọn đoạn. */
  fromOpenDoc: boolean;
  /** Việc người dùng cần tự làm (không sửa tự động được). */
  notes?: string[];
  /** Ghi chú phụ (chữ nhạt). */
  hints?: string[];
  /** Có kết quả này thì coi là đạt dù còn ghi chú. Mặc định chỉ xét lỗi định dạng. */
  maxIssues?: number;
}

const el = (tag: string, cls: string, text: string): HTMLElement => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  e.textContent = text;
  return e;
};

/** Vẽ kết quả kiểm tra vào `box`; trả về câu tóm tắt. */
export function renderCheckResult(box: HTMLElement, r: CheckResult, o: ViewOptions): string {
  const max = o.maxIssues ?? 300;
  box.replaceChildren();
  const clean = r.issues.length === 0 && r.page.length === 0 && r.tables.length === 0;
  const head = el('p', clean ? 'summary ok' : 'summary', clean ? o.okText : o.badText);
  box.append(head);

  for (const h of o.hints ?? []) box.append(el('p', 'hint', h));
  for (const pg of r.page) box.append(el('div', 'issue', `Trang: ${pg.field} là ${pg.actual} (chuẩn: ${pg.expected}). Chỉnh trong Bố cục > Lề.`));
  for (const ti of r.tables)
    box.append(
      el(
        'div',
        'issue',
        `Bảng ${ti.tableIndex + 1} (${ti.cols} cột): ${ti.diffs.map((x) => `${TABLE_FIELD_NAMES[x.field]} ${x.actual} (chuẩn: ${x.expected})`).join('; ')}`,
      ),
    );
  const canSelect = o.fromOpenDoc && !!o.host.selectParagraph;
  for (const i of r.issues.slice(0, max)) {
    const d = el(
      'div',
      canSelect ? 'issue go' : 'issue',
      `Đoạn ${i.index + 1} [${i.roleName}] “${i.text}”: ${i.diffs
        .map((x) => `${FIELD_NAMES[x.field]} ${showValue(x.field, x.actual)} (chuẩn: ${x.expectedLabel ?? showValue(x.field, x.expected)})`)
        .join('; ')}`,
    );
    if (canSelect) {
      d.title = 'Bấm để chọn đoạn này trong tài liệu';
      d.addEventListener('click', () => void o.host.selectParagraph?.(i.index));
    }
    box.append(d);
  }
  if (r.issues.length > max) box.append(el('p', 'hint', `… còn ${r.issues.length - max} đoạn khác.`));

  if (o.notes?.length) {
    box.append(el('h3', '', 'Cần bạn tự bổ sung hoặc sửa'));
    for (const n of o.notes) box.append(el('div', 'issue', n));
  }
  return head.textContent ?? '';
}
