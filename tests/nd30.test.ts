import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { joinHardWraps, normalizeText } from '../src/text';
import {
  assemble,
  classify,
  formatDate,
  formatNd30,
  parsePlaceDate,
  type Label,
} from '../src/profiles/vn-nd30';

const fx = (name: string): string => readFileSync(new URL(`../fixtures/${name}.txt`, import.meta.url), 'utf8');
const labelsOf = (raw: string): Label[] => classify(raw).map((l) => l.label);
const textOf = (raw: string, label: Label): string[] =>
  classify(raw)
    .filter((l) => l.label === label)
    .map((l) => l.text);

async function documentXml(docx: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(docx);
  return zip.file('word/document.xml')!.async('string');
}

describe('chuẩn hoá văn bản', () => {
  it('bỏ dòng trống và nối dòng bị ngắt cứng', () => {
    const lines = joinHardWraps(
      normalizeText('Phòng Nội vụ đề nghị các phòng, ban, đơn vị\n\nbáo cáo kết quả thực hiện.\nĐề nghị thực hiện.'),
    );
    expect(lines).toEqual([
      'Phòng Nội vụ đề nghị các phòng, ban, đơn vị báo cáo kết quả thực hiện.',
      'Đề nghị thực hiện.',
    ]);
  });
});

describe('ngày tháng', () => {
  it('thêm số 0 cho ngày < 10 và tháng 1, 2', () => {
    expect(formatDate({ place: 'Hà Nội', day: 5, month: 3, year: 2025 })).toBe('Hà Nội, ngày 05 tháng 3 năm 2025');
    expect(formatDate({ place: 'Hà Nội', day: 12, month: 1, year: 2025 })).toBe('Hà Nội, ngày 12 tháng 01 năm 2025');
    expect(formatDate({ place: 'Hà Nội', day: 1, month: 11, year: 2025 })).toBe('Hà Nội, ngày 01 tháng 11 năm 2025');
  });
  it('nhận diện các dạng địa danh, ngày tháng', () => {
    expect(parsePlaceDate('Hà Nội, ngày 5 tháng 3 năm 2025')).toEqual({ place: 'Hà Nội', day: 5, month: 3, year: 2025 });
    expect(parsePlaceDate('Thành phố Hồ Chí Minh, 05/03/2025')).toEqual({ place: 'Thành phố Hồ Chí Minh', day: 5, month: 3, year: 2025 });
    expect(parsePlaceDate('ngày … tháng … năm …')).toEqual({});
    expect(parsePlaceDate('Số liệu năm 2025 tăng 5 điểm')).toBeNull();
  });
});

describe('công văn', () => {
  const raw = fx('cong-van');
  it('nhận diện đủ các thành phần', () => {
    expect(textOf(raw, 'parent-org')).toEqual(['UBND HUYỆN MẪU']);
    expect(textOf(raw, 'issuer')).toEqual(['PHÒNG NỘI VỤ']);
    expect(textOf(raw, 'number')).toEqual(['Số: 45/PNV-VP']);
    expect(textOf(raw, 'subject-cv')).toHaveLength(1);
    expect(textOf(raw, 'addressee')).toEqual(['Các phòng, ban, đơn vị thuộc huyện']);
    expect(textOf(raw, 'recipient')).toEqual(['- Như trên;', '- Lưu: VT, NV.']);
    expect(textOf(raw, 'signer-title')).toEqual(['TRƯỞNG PHÒNG']);
    expect(textOf(raw, 'signer-name')).toEqual(['Nguyễn Văn An']);
    expect(textOf(raw, 'signature-note')).toEqual(['(đã ký)']);
  });
  it('nối dòng ngắt cứng trong nội dung', () => {
    const body = textOf(raw, 'paragraph')[0] ?? '';
    expect(body).toContain('đơn vị báo cáo kết quả');
  });
  it('không có tên loại thì là công văn', () => {
    const m = assemble(classify(raw));
    expect(m.docType).toBeUndefined();
    expect(m.vv).toMatch(/^đề nghị báo cáo/);
  });
});

describe('quyết định (dán từ bảng Word có tab)', () => {
  const raw = fx('quyet-dinh');
  it('tách cột phần đầu và sinh lại quốc hiệu, tiêu ngữ', () => {
    const labels = labelsOf(raw);
    expect(labels).toContain('national-title');
    expect(labels).toContain('motto');
    expect(textOf(raw, 'issuer')).toEqual(['ỦY BAN NHÂN DÂN HUYỆN MẪU']);
    expect(textOf(raw, 'place-date')).toEqual(['Mẫu, ngày 2 tháng 6 năm 2025']);
  });
  it('nhận diện tên loại, trích yếu, thẩm quyền, căn cứ, điều', () => {
    expect(textOf(raw, 'doc-type')).toEqual(['QUYẾT ĐỊNH']);
    expect(textOf(raw, 'subject')[0]).toMatch(/^Về việc ban hành/);
    expect(textOf(raw, 'authority')).toEqual(['CHỦ TỊCH ỦY BAN NHÂN DÂN HUYỆN MẪU']);
    // 2 dòng "Căn cứ" + dòng "Theo đề nghị của …," (cùng khối căn cứ, in nghiêng)
    expect(textOf(raw, 'basis')).toHaveLength(3);
    expect(textOf(raw, 'enact')).toEqual(['QUYẾT ĐỊNH:']);
    expect(textOf(raw, 'article')).toHaveLength(3);
  });
  it('chữ ký đặt sau danh sách nơi nhận, nhiều dòng chức vụ', () => {
    expect(textOf(raw, 'signer-title')).toEqual(['TM. ỦY BAN NHÂN DÂN', 'CHỦ TỊCH']);
    expect(textOf(raw, 'signer-name')).toEqual(['Trần Thị Bình']);
    expect(textOf(raw, 'recipient')).toHaveLength(3);
  });
});

describe('báo cáo', () => {
  const raw = fx('bao-cao');
  it('nhận diện tên loại, trích yếu, kính gửi, tiêu đề các cấp', () => {
    expect(textOf(raw, 'doc-type')).toEqual(['BÁO CÁO']);
    expect(textOf(raw, 'subject')[0]).toMatch(/^Kết quả thực hiện/);
    expect(textOf(raw, 'addressee')).toEqual(['Ủy ban nhân dân tỉnh ABC']);
    expect(textOf(raw, 'heading1')).toHaveLength(3);
    expect(textOf(raw, 'heading2')).toHaveLength(2);
  });
  it('chữ ký đặt trước nơi nhận', () => {
    expect(textOf(raw, 'signer-title')).toEqual(['GIÁM ĐỐC']);
    expect(textOf(raw, 'signer-name')).toEqual(['Lê Hoàng Nam']);
  });
  it('nhận diện cơ quan ban hành khi chỉ có một dòng tên cơ quan', () => {
    const issuer = classify(raw).find((l) => l.label === 'issuer');
    expect(issuer?.text).toBe('SỞ XÂY DỰNG TỈNH ABC');
  });
});

describe('tên loại và trích yếu trên cùng một dòng', () => {
  it('tách "BÁO CÁO TỔNG KẾT NĂM 2025" thành tên loại + trích yếu', () => {
    const lines = classify('SỞ NỘI VỤ\nSố: 1/BC-SNV\nBÁO CÁO TỔNG KẾT NĂM 2025\nNội dung báo cáo.');
    expect(lines.find((l) => l.label === 'doc-type')?.text).toBe('BÁO CÁO');
    expect(lines.find((l) => l.label === 'subject')?.text).toBe('Tổng kết năm 2025');
  });
});

describe('xuất .docx', () => {
  it('đúng khổ giấy, lề, font, đánh số trang và nội dung', async () => {
    const { docx } = await formatNd30(fx('quyet-dinh'));
    const xml = await documentXml(docx);
    // A4 = 11906 x 16838 twips; lề 20/20/30/20 mm
    expect(xml).toContain('w:w="11906"');
    expect(xml).toContain('w:h="16838"');
    expect(xml).toMatch(/w:top="1134"/);
    expect(xml).toMatch(/w:left="1701"/);
    expect(xml).toMatch(/w:right="1134"/);
    expect(xml).toContain('w:titlePg');
    expect(xml).toContain('Times New Roman');
    expect(xml).toContain('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM');
    expect(xml).toContain('Độc lập - Tự do - Hạnh phúc');
    expect(xml).toContain('Mẫu, ngày 02 tháng 6 năm 2025');
    expect(xml).toContain('Nơi nhận:');
    expect(xml).toContain('- Lưu: VT, VP.');
    expect(xml).toContain('Trần Thị Bình');
  });

  it('tùy chọn của người dùng ghi đè và bổ sung thông tin còn thiếu', async () => {
    const { model, docx } = await formatNd30('Kính gửi: Sở A\nNội dung công văn đề nghị phối hợp xử lý.', {
      issuer: 'SỞ TÀI CHÍNH',
      place: 'Hà Nội',
      dateIso: '2025-01-09',
      number: '9/STC-NS',
      signerTitle: 'GIÁM ĐỐC',
      signerName: 'Phạm Văn Cường',
    });
    expect(model.issuer).toBe('SỞ TÀI CHÍNH');
    const xml = await documentXml(docx);
    expect(xml).toContain('Hà Nội, ngày 09 tháng 01 năm 2025');
    expect(xml).toContain('Số: 9/STC-NS');
    expect(xml).toContain('Phạm Văn Cường');
    expect(model.warnings.join(' ')).not.toContain('Thiếu số');
  });

  it('văn bản thiếu thông tin vẫn tạo được .docx và có cảnh báo', async () => {
    const { model, docx } = await formatNd30('Đây là một đoạn văn bản thô duy nhất.');
    expect(docx.byteLength).toBeGreaterThan(1000);
    expect(model.warnings.length).toBeGreaterThan(3);
  });
});
