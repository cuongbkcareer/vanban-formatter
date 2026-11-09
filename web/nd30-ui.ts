import { readDocx, type DocFormat } from '../src/docx-read';
import { checkNd30, planNd30Fixes } from '../src/profiles/vn-nd30/check';
import { renderCheckResult } from './check-view';
import type { Host } from './ui';

type Status = (msg: string, kind?: 'ok' | 'error') => void;

/** Kiểm tra và sửa định dạng một văn bản đã soạn theo NĐ 30/2020 (chạy trên máy, không dùng AI). */
export function mountNd30Check(host: Host, status: Status): void {
  const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;
  const box = $('n-result');

  const show = (doc: DocFormat, fromOpenDoc: boolean): string => {
    const r = checkNd30(doc);
    const text = renderCheckResult(box, r, {
      okText: `Đúng thể thức: ${r.paragraphs} đoạn đều đúng định dạng.`,
      badText: `${r.issues.length}/${r.paragraphs} đoạn lệch định dạng, ${r.tables.length} bảng sai khung, ${r.page.length} lỗi trang.`,
      host,
      fromOpenDoc,
      notes: r.notes,
    });
    status(text);
    return text;
  };

  const fail = (what: string, e: unknown): void => status(`${what}: ${e instanceof Error ? e.message : String(e)}`, 'error');

  $('n-check-doc').addEventListener('click', async () => {
    try {
      if (host.readFormat) show(await host.readFormat(), true);
    } catch (e) {
      fail('Không đọc được tài liệu đang mở', e);
    }
  });

  const file = $<HTMLInputElement>('n-file');
  $('n-check-file').addEventListener('click', () => file.click());
  file.addEventListener('change', async () => {
    const f = file.files?.[0];
    file.value = '';
    if (!f) return;
    try {
      show(await readDocx(await f.arrayBuffer()), false);
    } catch (e) {
      fail('Không đọc được file', e);
    }
  });

  $('n-fix').addEventListener('click', async () => {
    if (!host.readFormat || !host.applyFixes) return;
    const btn = $<HTMLButtonElement>('n-fix');
    btn.disabled = true;
    try {
      const doc = await host.readFormat();
      const plan = planNd30Fixes(checkNd30(doc));
      if (plan.items.length === 0 && plan.tableItems.length === 0) {
        show(doc, true);
        return;
      }
      const res = plan.items.length ? await host.applyFixes(plan.items) : { applied: 0, skipped: 0, capsNotApplied: 0 };
      // Sửa bảng sau cùng: đổi viền có thể làm thay đổi bố cục các đoạn trong bảng.
      const tres = plan.tableItems.length && host.applyTableFixes ? await host.applyTableFixes(plan.tableItems) : { applied: 0, notes: [] as string[] };
      const after = show(await host.readFormat(), true);
      const notes = [
        `Đã sửa ${res.applied} đoạn${plan.tableItems.length ? ` và ${tres.applied}/${plan.tableItems.length} bảng` : ''} theo NĐ 30.`,
        ...tres.notes,
        res.skipped ? `Bỏ qua ${res.skipped} đoạn (nội dung đã thay đổi hoặc không khớp).` : '',
        res.capsNotApplied ? `${res.capsNotApplied} đoạn cần chữ in hoa: hãy tự gõ hoa (công cụ không đổi chữ của bạn).` : '',
        'Có thể hoàn tác bằng Cmd+Z / Ctrl+Z.',
        `Kiểm tra lại: ${after}`,
      ].filter(Boolean);
      status(notes.join(' '));
    } catch (e) {
      fail('Không sửa được', e);
    } finally {
      btn.disabled = false;
    }
  });
}
