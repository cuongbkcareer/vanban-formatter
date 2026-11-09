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
import { renderDocxBase64 } from '../src/render-docx';

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

const OPTION_IDS = ['o-type', 'o-parent', 'o-issuer', 'o-place', 'o-date', 'o-number', 'o-stitle', 'o-sname'];

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

  const analyze = (): void => {
    const raw = $<HTMLTextAreaElement>('raw').value;
    if (!raw.trim()) {
      status('Chưa có văn bản để phân tích.', 'error');
      return;
    }
    lines = classify(raw);
    $('review').classList.remove('hidden');
    renderLines();
    renderWarnings();
    status(`Đã nhận diện ${lines.length} dòng. Hãy kiểm tra các dòng nền vàng.`);
  };

  const makeDocx = async (): Promise<{ base64: string; filename: string }> => {
    const model = assemble(lines, readOptions());
    const base64 = await renderDocxBase64(buildIR(model));
    const filename = `${(model.docType ?? CONG_VAN).replace(/\s+/g, '-').toLowerCase()}.docx`;
    return { base64, filename };
  };

  const act = (fn: (d: { base64: string; filename: string }) => Promise<void> | void, okMsg: string) => async (): Promise<void> => {
    try {
      if (lines.length === 0) analyze();
      if (lines.length === 0) return;
      await fn(await makeDocx());
      status(okMsg);
    } catch (e) {
      status(`Lỗi: ${e instanceof Error ? e.message : String(e)}`, 'error');
    }
  };

  const typeSel = $<HTMLSelectElement>('o-type');
  for (const t of [CONG_VAN, ...DOC_TYPES]) {
    const opt = document.createElement('option');
    opt.value = t;
    opt.textContent = t;
    typeSel.append(opt);
  }

  $('analyze').addEventListener('click', analyze);
  $('sample').addEventListener('click', () => {
    $<HTMLTextAreaElement>('raw').value = SAMPLE;
  });
  $('read-doc').addEventListener('click', async () => {
    try {
      const text = await host.readSource?.();
      if (!text?.trim()) {
        status('Tài liệu đang trống.', 'error');
        return;
      }
      $<HTMLTextAreaElement>('raw').value = text;
      analyze();
    } catch (e) {
      status(`Không đọc được tài liệu: ${e instanceof Error ? e.message : String(e)}`, 'error');
    }
  });
  $('download').addEventListener('click', act(({ base64, filename }) => host.save?.(base64, filename), 'Đã tạo file .docx.'));
  $('open-new').addEventListener('click', act(({ base64 }) => host.openNew?.(base64), 'Đã mở văn bản đúng thể thức thành tài liệu mới.'));
  $('replace').addEventListener('click', act(({ base64 }) => host.replaceCurrent?.(base64), 'Đã thay nội dung. Hãy kiểm tra lề trang và số trang.'));
  for (const id of OPTION_IDS) $(id).addEventListener('input', () => lines.length && renderWarnings());
}
