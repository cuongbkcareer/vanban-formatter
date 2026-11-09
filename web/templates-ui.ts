import { readDocx, type DocFormat } from '../src/docx-read';
import {
  checkAgainstProfile,
  describeTableSpec,
  learnProfile,
  planFixes,
  planTableFixes,
  type FormatProfile,
  type Role,
} from '../src/learn/profile';
import { renderCheckResult } from './check-view';
import { getActiveName, isProfile, loadProfiles, saveProfiles, setActiveName } from './storage';
import type { Host } from './ui';

type Status = (msg: string, kind?: 'ok' | 'error') => void;

const ALIGN: Record<string, string> = { left: 'trái', center: 'giữa', right: 'phải', justify: 'đều' };

function summarize(r: Role): string {
  const f = r.format;
  const parts = [
    `${r.count} đoạn`,
    `căn ${ALIGN[f.align] ?? f.align}`,
    `${f.sizePt} pt`,
    f.bold ? 'đậm' : '',
    f.italic ? 'nghiêng' : '',
    f.caps ? 'in hoa' : '',
    f.firstLineMm ? `thụt đầu dòng ${f.firstLineMm} mm` : '',
    f.leftMm ? `lề trái +${f.leftMm} mm` : '',
    f.spaceBeforePt ? `cách trước ${f.spaceBeforePt} pt` : '',
    f.lineSpacing && f.lineSpacing !== 1 ? `giãn dòng ×${f.lineSpacing}` : '',
  ].filter(Boolean);
  return `${parts.join(' · ')}. Ví dụ: ${r.examples.map((e) => `“${e}”`).join(' | ')}`;
}

export function mountTemplates(host: Host, status: Status): void {
  const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;
  let profiles = loadProfiles();
  let learned: FormatProfile | null = null;

  const select = $<HTMLSelectElement>('t-active');
  const active = (): FormatProfile | null => profiles.find((p) => p.name === select.value) ?? null;

  function refreshSelect(prefer?: string): void {
    select.replaceChildren();
    const none = document.createElement('option');
    none.value = '';
    none.textContent = profiles.length ? '(chọn một mẫu)' : '(chưa có mẫu)';
    select.append(none);
    for (const p of profiles) {
      const o = document.createElement('option');
      o.value = p.name;
      o.textContent = p.name;
      select.append(o);
    }
    const want = prefer ?? getActiveName();
    select.value = profiles.some((p) => p.name === want) ? want : '';
  }
  refreshSelect();
  select.addEventListener('change', () => setActiveName(select.value));

  const fail = (what: string, e: unknown): void => status(`${what}: ${e instanceof Error ? e.message : String(e)}`, 'error');

  /* ---------------- Học từ file mẫu ---------------- */
  function showLearned(p: FormatProfile): void {
    learned = p;
    $('t-learned').classList.remove('hidden');
    $<HTMLInputElement>('t-name').value = p.name;
    const box = $('t-roles');
    box.replaceChildren();
    for (const r of p.roles) {
      const div = document.createElement('div');
      div.className = 'role';
      const input = document.createElement('input');
      input.value = r.name;
      input.setAttribute('aria-label', 'Tên vai trò');
      input.addEventListener('input', () => {
        r.name = input.value;
      });
      const info = document.createElement('div');
      info.className = 'info';
      info.textContent = summarize(r);
      div.append(input, info);
      box.append(div);
    }
    for (const spec of p.tables ?? []) {
      const div = document.createElement('div');
      div.className = 'role';
      div.textContent = describeTableSpec(spec);
      box.append(div);
    }
    status(`Đã học ${p.roles.length} vai trò${p.tables?.length ? ` và ${p.tables.length} khung bảng` : ''} từ ${p.createdFrom ?? 'tài liệu mẫu'}. Đổi tên cho dễ nhớ rồi bấm Lưu mẫu.`);
  }

  async function learnFrom(doc: DocFormat, label: string): Promise<void> {
    const base = label.replace(/\.docx$/i, '');
    showLearned(learnProfile(doc, base, label));
  }

  const fileInput = (id: string, onFile: (f: File) => Promise<void>): void => {
    const el = $<HTMLInputElement>(id);
    el.addEventListener('change', async () => {
      const f = el.files?.[0];
      el.value = '';
      if (!f) return;
      try {
        await onFile(f);
      } catch (e) {
        fail('Không đọc được file', e);
      }
    });
  };
  fileInput('t-file-docx', async (f) => learnFrom(await readDocx(await f.arrayBuffer()), f.name));
  $('t-learn-file').addEventListener('click', () => $('t-file-docx').click());
  $('t-learn-doc').addEventListener('click', async () => {
    try {
      if (!host.readFormat) return;
      await learnFrom(await host.readFormat(), 'tài liệu đang mở');
    } catch (e) {
      fail('Không đọc được định dạng của tài liệu đang mở', e);
    }
  });

  $('t-save').addEventListener('click', () => {
    if (!learned) return;
    const name = $<HTMLInputElement>('t-name').value.trim() || learned.name;
    learned.name = name;
    profiles = [...profiles.filter((p) => p.name !== name), learned];
    saveProfiles(profiles);
    setActiveName(name);
    refreshSelect(name);
    $('t-learned').classList.add('hidden');
    learned = null;
    status(`Đã lưu mẫu “${name}” và chọn làm mẫu đang dùng.`);
  });

  /* ---------------- Xuất, nhập, xoá ---------------- */
  $('t-export').addEventListener('click', () => {
    const p = active();
    if (!p) return status('Hãy chọn một mẫu để xuất.', 'error');
    const json = JSON.stringify(p, null, 2);
    const area = $<HTMLTextAreaElement>('t-json');
    area.value = json;
    area.classList.remove('hidden');
    try {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([json], { type: 'application/json' }));
      a.download = `mau-${p.name.replace(/[^\p{L}\p{N}]+/gu, '-')}.json`;
      document.body.append(a);
      a.click();
      a.remove();
    } catch {
      /* khung add-in có thể chặn tải về: người dùng sao chép từ ô văn bản */
    }
    status('Đã xuất mẫu. Nếu không thấy file tải về, hãy sao chép nội dung trong ô bên dưới.');
  });
  fileInput('t-file-json', async (f) => {
    const data: unknown = JSON.parse(await f.text());
    if (!isProfile(data)) throw new Error('File không phải mẫu hợp lệ.');
    let name = data.name;
    for (let i = 2; profiles.some((p) => p.name === name); i++) name = `${data.name} (${i})`;
    data.name = name;
    profiles = [...profiles, data];
    saveProfiles(profiles);
    setActiveName(name);
    refreshSelect(name);
    status(`Đã nhập mẫu “${name}”.`);
  });
  $('t-import').addEventListener('click', () => $('t-file-json').click());
  $('t-delete').addEventListener('click', () => {
    const p = active();
    if (!p) return status('Hãy chọn một mẫu để xoá.', 'error');
    profiles = profiles.filter((x) => x.name !== p.name);
    saveProfiles(profiles);
    setActiveName('');
    refreshSelect('');
    status(`Đã xoá mẫu “${p.name}”.`);
  });

  /* ---------------- Kiểm tra ---------------- */
  const result = $('t-result');
  function runCheck(doc: DocFormat, fromOpenDoc: boolean): void {
    const profile = active();
    if (!profile) return status('Hãy chọn hoặc học một mẫu trước khi kiểm tra.', 'error');
    const r = checkAgainstProfile(profile, doc);
    const hints: string[] = [];
    if (!doc.page) hints.push('Không đọc được lề trang từ nguồn này nên chưa kiểm tra lề.');
    if (r.unmatchedTables.length)
      hints.push(
        `${r.unmatchedTables.length} bảng có số cột khác với bảng trong mẫu nên chưa kiểm tra khung (bảng ${r.unmatchedTables.map((x) => x + 1).join(', ')}). Muốn kiểm tra, hãy thêm một bảng cùng số cột vào file mẫu rồi học lại.`,
      );
    const text = renderCheckResult(result, r, {
      okText: `Đạt mẫu “${profile.name}”: ${r.paragraphs} đoạn đều đúng định dạng.`,
      badText: `Theo mẫu “${profile.name}”: ${r.issues.length}/${r.paragraphs} đoạn lệch định dạng, ${r.tables.length} bảng sai khung, ${r.page.length} lỗi trang.`,
      host,
      fromOpenDoc,
      hints,
    });
    status(text);
  }

  $('t-check-doc').addEventListener('click', async () => {
    try {
      if (host.readFormat) runCheck(await host.readFormat(), true);
    } catch (e) {
      fail('Không đọc được tài liệu đang mở', e);
    }
  });
  fileInput('t-file-check', async (f) => runCheck(await readDocx(await f.arrayBuffer()), false));
  $('t-check-file').addEventListener('click', () => $('t-file-check').click());

  /* ---------------- Sửa tự động (Word) ---------------- */
  $('t-fix').addEventListener('click', async () => {
    const profile = active();
    if (!profile) return status('Hãy chọn hoặc học một mẫu trước.', 'error');
    if (!host.readFormat || !host.applyFixes) return;
    try {
      const doc = await host.readFormat();
      const items = planFixes(profile, doc);
      const tableItems = planTableFixes(profile, doc);
      if (items.length === 0 && tableItems.length === 0) {
        runCheck(doc, true);
        return;
      }
      const res = items.length ? await host.applyFixes(items, profile) : { applied: 0, skipped: 0, capsNotApplied: 0 };
      // Sửa bảng sau cùng: đổi độ rộng/viền có thể làm thay đổi bố cục các đoạn trong bảng.
      const tres = tableItems.length && host.applyTableFixes ? await host.applyTableFixes(tableItems) : { applied: 0, notes: [] as string[] };
      const notes = [
        `Đã sửa ${res.applied} đoạn${tableItems.length ? ` và ${tres.applied}/${tableItems.length} bảng` : ''} theo mẫu “${profile.name}”.`,
        ...tres.notes,
        res.skipped ? `Bỏ qua ${res.skipped} đoạn (nội dung đã thay đổi hoặc không khớp).` : '',
        res.capsNotApplied ? `${res.capsNotApplied} đoạn cần chữ in hoa: hãy tự gõ hoa (công cụ không đổi chữ của bạn).` : '',
        'Có thể hoàn tác bằng Cmd+Z / Ctrl+Z.',
      ].filter(Boolean);
      runCheck(await host.readFormat(), true);
      status(notes.join(' '));
    } catch (e) {
      fail('Không sửa được', e);
    }
  });
}
