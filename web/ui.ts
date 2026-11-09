import {
  assemble,
  buildIR,
  classify,
  CONG_VAN,
  DOC_TYPES,
  LABEL_NAMES,
  type Label,
  type LabeledLine,
  type Options,
} from '../src/profiles/vn-nd30';
import type { DocFormat } from '../src/docx-read';
import type { FixItem, FormatProfile, TableFix } from '../src/learn/profile';
import { renderDocxBase64 } from '../src/render-docx';
import { mountNd30Check } from './nd30-ui';
import { mountTemplates } from './templates-ui';
import { loadOptions, ORG_FIELD_IDS, saveOptions } from './storage';

/** Nơi giao diện chạy: trang web độc lập hoặc khung bên cạnh trong Word. */
export interface Host {
  kind: 'web' | 'word';
  /** Đọc văn bản từ tài liệu đang mở (vùng chọn, nếu không có thì toàn bộ). Chỉ có trong Word. */
  readSource?: () => Promise<string>;
  /** Tải file về (trang web). */
  save?: (base64: string, filename: string) => Promise<void> | void;
  /** Mở thành tài liệu Word mới. */
  openNew?: (base64: string) => Promise<void>;
  /** Thay nội dung tài liệu đang mở. */
  replaceCurrent?: (base64: string) => Promise<void>;
  /** Đọc định dạng của tài liệu đang mở (chỉ có trong Word). */
  readFormat?: () => Promise<DocFormat>;
  /** Chọn (bôi đen) một đoạn trong tài liệu đang mở theo chỉ số đoạn. */
  selectParagraph?: (index: number) => Promise<void>;
  /** Sửa định dạng trực tiếp trong tài liệu đang mở. */
  applyFixes?: (items: FixItem[], profile?: FormatProfile) => Promise<{ applied: number; skipped: number; capsNotApplied: number }>;
  /** Sửa khung bảng (viền, độ rộng cột, tô nền hàng tiêu đề ...) trong tài liệu đang mở. */
  applyTableFixes?: (items: TableFix[]) => Promise<{ applied: number; notes: string[] }>;
}

const SAMPLE = `UBND HUYỆN MẪU
PHÒNG NỘI VỤ
Số: 45/PNV-VP
V/v đề nghị báo cáo kết quả cải cách hành chính năm 2024
Mẫu, ngày 5 tháng 3 năm 2025
Kính gửi: Các phòng, ban, đơn vị thuộc huyện
Thực hiện Kế hoạch số 12/KH-UBND của Ủy ban nhân dân huyện, Phòng Nội vụ đề nghị các đơn vị báo cáo kết quả thực hiện công tác cải cách hành chính năm 2024.
Báo cáo gửi về Phòng Nội vụ trước ngày 20 tháng 3 năm 2025.
Nơi nhận:
- Như trên;
- Lưu: VT, NV.
TRƯỞNG PHÒNG
Nguyễn Văn An`;

export function mountUI(host: Host): void {
  const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;
  document.body.dataset.host = host.kind;
  let lines: LabeledLine[] = [];

  const status = (msg: string, kind: 'ok' | 'error' = 'ok'): void => {
    const el = $('status');
    el.textContent = msg;
    el.dataset.kind = kind;
  };

  const readOptions = (): Options => {
    const o: Options = {};
    const v = (id: string): string => $<HTMLInputElement>(id).value.trim();
    if (v('o-type')) o.docType = v('o-type');
    if (v('o-parent')) o.parentOrg = v('o-parent');
    if (v('o-issuer')) o.issuer = v('o-issuer');
    if (v('o-place')) o.place = v('o-place');
    if (v('o-date')) o.dateIso = v('o-date');
    if (v('o-number')) o.number = v('o-number');
    if (v('o-stitle')) o.signerTitle = v('o-stitle').split('|').map((s) => s.trim()).join('\n');
    if (v('o-sname')) o.signerName = v('o-sname');
    return o;
  };

  const renderWarnings = (): void => {
    const ul = $('warnings');
    ul.replaceChildren();
    for (const w of assemble(lines, readOptions()).warnings) {
      const li = document.createElement('li');
      li.textContent = w;
      ul.append(li);
    }
  };

  const renderLines = (): void => {
    const box = $('lines');
    box.replaceChildren();
    lines.forEach((line, idx) => {
      const row = document.createElement('div');
      row.className = line.confidence < 0.7 ? 'line low' : 'line';

      const sel = document.createElement('select');
      sel.className = 'sel';
      sel.setAttribute('aria-label', 'Nhãn của dòng');
      for (const [label, name] of Object.entries(LABEL_NAMES)) {
        const opt = document.createElement('option');
        opt.value = label;
        opt.textContent = name;
        sel.append(opt);
      }
      sel.value = line.label;
      sel.addEventListener('change', () => {
        line.label = sel.value as Label;
        line.confidence = 1;
        row.classList.remove('low');
        conf.textContent = '1.00';
        renderWarnings();
      });

      const input = document.createElement('input');
      input.className = 'txt';
      input.setAttribute('aria-label', 'Nội dung dòng');
      input.value = line.text;
      input.addEventListener('input', () => {
        line.text = input.value;
        renderWarnings();
      });

      const conf = document.createElement('span');
      conf.className = 'conf';
      conf.title = 'Độ tin cậy nhận diện';
      conf.textContent = line.confidence.toFixed(2);

      const del = document.createElement('button');
      del.className = 'del';
      del.textContent = '×';
      del.title = 'Xoá dòng';
      del.setAttribute('aria-label', 'Xoá dòng');
      del.addEventListener('click', () => {
        lines.splice(idx, 1);
        renderLines();
        renderWarnings();
      });

      row.append(sel, input, conf, del);
      box.append(row);
    });
  };

  const getRaw = async (): Promise<string> =>
    host.readSource ? host.readSource() : $<HTMLTextAreaElement>('raw').value;

  const showLines = (): void => {
    $('review').classList.remove('hidden');
    renderLines();
    renderWarnings();
  };

  /** Phân tích văn bản; trả về false nếu không có gì để xử lý. */
  const analyze = async (): Promise<boolean> => {
    const raw = await getRaw();
    if (!raw.trim()) {
      status(host.kind === 'word' ? 'Tài liệu đang trống.' : 'Chưa có văn bản để xử lý. Hãy dán văn bản vào ô trên.', 'error');
      return false;
    }
    lines = classify(raw);
    renderWarnings();
    if (!$('review').classList.contains('hidden')) renderLines();
    return true;
  };

  const makeDocx = async (): Promise<{ base64: string; filename: string }> => {
    const model = assemble(lines, readOptions());
    const base64 = await renderDocxBase64(buildIR(model));
    const filename = `${(model.docType ?? CONG_VAN).replace(/\s+/g, '-').toLowerCase()}.docx`;
    return { base64, filename };
  };

  /** Các nút trong danh sách kiểm tra: dùng đúng danh sách nhãn người dùng đã sửa. */
  const act = (fn: (d: { base64: string; filename: string }) => Promise<void> | void, okMsg: string) => async (): Promise<void> => {
    try {
      if (lines.length === 0 && !(await analyze())) return;
      await fn(await makeDocx());
      status(okMsg);
    } catch (e) {
      status(`Lỗi: ${e instanceof Error ? e.message : String(e)}`, 'error');
    }
  };

  /** Nút chính: đọc, nhận diện, tạo văn bản đúng thể thức, giao cho Word (hoặc tải về) trong một lần bấm. */
  const quick = async (): Promise<void> => {
    const btn = $<HTMLButtonElement>('quick');
    btn.disabled = true;
    status('Đang xử lý…');
    try {
      if (!(await analyze())) return;
      const { base64, filename } = await makeDocx();
      if (host.openNew) await host.openNew(base64);
      else await host.save?.(base64, filename);
      const missing = assemble(lines, readOptions()).warnings.filter((w) => w.startsWith('Thiếu') || w.startsWith('Chưa'));
      status(
        (host.kind === 'word' ? 'Đã mở văn bản đúng thể thức thành tài liệu mới.' : 'Đã tạo file .docx.') +
          (missing.length ? ` Còn ${missing.length} thông tin cần bổ sung (chỗ có dấu …), xem bên dưới.` : ''),
      );
    } catch (e) {
      status(`Lỗi: ${e instanceof Error ? e.message : String(e)}`, 'error');
    } finally {
      btn.disabled = false;
    }
  };

  const typeSel = $<HTMLSelectElement>('o-type');
  for (const t of [CONG_VAN, ...DOC_TYPES]) {
    const opt = document.createElement('option');
    opt.value = t;
    opt.textContent = t;
    typeSel.append(opt);
  }

  $('quick').textContent = host.kind === 'word' ? 'Chuẩn hoá văn bản này' : 'Tạo file .docx';
  $('quick').addEventListener('click', () => void quick());
  $('analyze').addEventListener('click', async () => {
    try {
      if (await analyze()) {
        showLines();
        status(`Đã nhận diện ${lines.length} dòng. Hãy kiểm tra các dòng nền vàng.`);
      }
    } catch (e) {
      status(`Không đọc được tài liệu: ${e instanceof Error ? e.message : String(e)}`, 'error');
    }
  });
  $('sample').addEventListener('click', () => {
    $<HTMLTextAreaElement>('raw').value = SAMPLE;
  });
  $('download').addEventListener('click', act(({ base64, filename }) => host.save?.(base64, filename), 'Đã tạo file .docx.'));
  $('open-new').addEventListener('click', act(({ base64 }) => host.openNew?.(base64), 'Đã mở văn bản đúng thể thức thành tài liệu mới.'));
  $('replace').addEventListener('click', act(({ base64 }) => host.replaceCurrent?.(base64), 'Đã thay nội dung. Hãy kiểm tra lề trang và số trang.'));
  // Nhớ thông tin cơ quan trên máy này (không gửi đi đâu). Số, ngày và loại văn bản thì không nhớ.
  const saved = loadOptions();
  const orgValues = (): Record<string, string> =>
    Object.fromEntries(ORG_FIELD_IDS.map((k) => [k, $<HTMLInputElement>(k).value]));
  for (const id of ORG_FIELD_IDS) {
    const el = $<HTMLInputElement>(id);
    if (saved[id]) el.value = saved[id];
    el.addEventListener('input', () => {
      saveOptions(orgValues());
      if (lines.length) renderWarnings();
    });
  }
  $('o-forget').addEventListener('click', () => {
    saveOptions({});
    for (const id of ORG_FIELD_IDS) $<HTMLInputElement>(id).value = '';
    status('Đã xoá thông tin cơ quan đã nhớ.');
  });
  mountTemplates(host, status);
  mountNd30Check(host, status);

  // Chuyển mục: "Chuẩn hoá" và "Mẫu định dạng".
  const tabIds = ['format', 'template'];
  const showTab = (name: string): void => {
    for (const t of tabIds) {
      $(`tab-${t}`).setAttribute('aria-selected', String(t === name));
      $(`pane-${t}`).hidden = t !== name;
    }
  };
  for (const t of tabIds) $(`tab-${t}`).addEventListener('click', () => showTab(t));
}
