import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { TEST_DOCS } from '../fixtures/test-docs';
import { assemble, classify, formatNd30, splitAgencyName, type Label } from '../src/profiles/vn-nd30';

const doc = (prefix: string) => {
  const d = TEST_DOCS.find((x) => x.file.startsWith(prefix));
  if (!d) throw new Error(`không có văn bản ${prefix}`);
  // Giống cách add-in lấy văn bản từ Word: các đoạn nối bằng \n.
  const lines = classify(d.lines.join('\n'));
  return { lines, model: assemble(lines) };
};
const texts = (lines: ReturnType<typeof classify>, label: Label) =>
  lines.filter((l) => l.label === label).map((l) => l.text);

describe('bộ văn bản thử add-in', () => {
  it('01 công văn sạch: không cảnh báo, tên cơ quan "SỞ …" được nhận diện chắc chắn', () => {
    const { lines, model } = doc('01');
    expect(model.warnings).toEqual([]);
    const issuer = lines.find((l) => l.label === 'issuer');
    expect(issuer?.text).toBe('SỞ GIÁO DỤC VÀ ĐÀO TẠO');
    // lỗi cũ: \b không khớp sau chữ có dấu ("Sở"), làm độ tin cậy tụt xuống 0.6
    expect(issuer?.confidence).toBeGreaterThanOrEqual(0.7);
    expect(model.parentOrg).toEqual(['UBND THÀNH PHỐ MẪU']);
    expect(model.signerTitle).toEqual(['KT. GIÁM ĐỐC', 'PHÓ GIÁM ĐỐC']);
  });

  it('02 "ỦY BAN NHÂN DÂN" + "THÀNH PHỐ MẪU" là một cơ quan ban hành', () => {
    const { model } = doc('02');
    expect(model.parentOrg).toEqual([]);
    expect(model.issuer).toBe('ỦY BAN NHÂN DÂN THÀNH PHỐ MẪU');
    expect(model.warnings).toEqual([]);
  });

  it('02 dòng "Theo đề nghị của…," thuộc phần căn cứ (in nghiêng như các dòng Căn cứ)', () => {
    const { lines } = doc('02');
    expect(texts(lines, 'basis')).toHaveLength(3);
    expect(texts(lines, 'basis')[2]).toMatch(/^Theo đề nghị của/);
    expect(texts(lines, 'enact')).toEqual(['QUYẾT ĐỊNH:']);
  });

  it('chỉ nhận "Theo/Xét …" là căn cứ khi nằm trong khối căn cứ, không phải đoạn nội dung sau "QUYẾT ĐỊNH:"', () => {
    const lines = classify(['QUYẾT ĐỊNH', 'Về việc A', 'Căn cứ Luật X;', 'Theo đề nghị của Sở Y,', 'QUYẾT ĐỊNH:', 'Điều 1. Z.', 'Theo quy định này, các đơn vị thực hiện;'].join('\n'));
    expect(lines.filter((l) => l.label === 'basis')).toHaveLength(2);
    expect(lines.find((l) => l.text.startsWith('Theo quy định'))?.label).toBe('paragraph');
  });

  it('ngắt tên cơ quan đúng chỗ khi in', () => {
    expect(splitAgencyName('ỦY BAN NHÂN DÂN THÀNH PHỐ MẪU')).toEqual(['ỦY BAN NHÂN DÂN', 'THÀNH PHỐ MẪU']);
    expect(splitAgencyName('HỘI ĐỒNG NHÂN DÂN TỈNH ABC')).toEqual(['HỘI ĐỒNG NHÂN DÂN', 'TỈNH ABC']);
    expect(splitAgencyName('SỞ XÂY DỰNG TỈNH ABC')).toEqual(['SỞ XÂY DỰNG TỈNH ABC']);
    expect(splitAgencyName('ỦY BAN NHÂN DÂN')).toEqual(['ỦY BAN NHÂN DÂN']);
  });

  it('03 quy chế: chương, điều, khoản', () => {
    const { lines } = doc('03');
    expect(texts(lines, 'chapter')).toEqual(['Chương I', 'Chương II']);
    expect(texts(lines, 'chapter-title')).toEqual(['QUY ĐỊNH CHUNG', 'TỔ CHỨC THỰC HIỆN']);
    expect(texts(lines, 'article')).toHaveLength(4);
    expect(lines.some((l) => l.label.startsWith('heading'))).toBe(false);
    expect(texts(lines, 'attachment-note')).toEqual(['(Ban hành kèm theo Quyết định số 56/QĐ-UBND ngày 12 tháng 1 năm 2026)']);
  });

  it('04 báo cáo: các cấp tiêu đề', () => {
    const { lines, model } = doc('04');
    expect(texts(lines, 'heading1')).toHaveLength(3);
    expect(texts(lines, 'heading2')).toHaveLength(2);
    expect(model.warnings).toEqual([]);
  });

  it('05 tờ trình: kính gửi nhiều nơi và chữ ký KT.', () => {
    const { model } = doc('05');
    expect(model.addressees).toHaveLength(2);
    expect(model.signerTitle).toEqual(['KT. TRƯỞNG PHÒNG', 'PHÓ TRƯỞNG PHÒNG']);
    expect(model.warnings).toEqual([]);
  });

  it('06 thông báo: đủ thành phần', () => {
    const { model } = doc('06');
    expect(model.docType).toBe('Thông báo');
    expect(model.recipients).toHaveLength(3);
    expect(model.warnings).toEqual([]);
  });

  it('07 lộn xộn: tách cột bằng Tab, gộp dòng ngắt, nơi nhận sau chữ ký', () => {
    const { lines, model } = doc('07');
    expect(model.issuer).toBe('ỦY BAN NHÂN DÂN XÃ MẪU');
    expect(model.number).toBe('12/UBND-VP');
    expect(model.placeDate).toEqual({ place: 'Mẫu', day: 5, month: 3, year: 2026 });
    expect(model.signerName).toBe('NGUYỄN VĂN HÙNG');
    expect(model.recipients).toEqual(['Như trên;', 'Lưu: VT.']);
    expect(texts(lines, 'paragraph').some((t) => t.includes('phòng chống thiên tai năm 2026 vào lúc 14 giờ'))).toBe(true);
    expect(model.warnings).toEqual([]);
  });

  it('08 thiếu thông tin: cảnh báo đủ các mục còn thiếu nhưng vẫn dựng được', () => {
    const { model } = doc('08');
    const w = model.warnings.join(' | ');
    for (const part of ['cơ quan ban hành', 'số, ký hiệu', 'địa danh', 'ngày, tháng, năm', 'chức vụ người ký', 'họ tên người ký', 'nơi nhận']) {
      expect(w.toLowerCase()).toContain(part);
    }
  });
});

describe('định dạng trong file .docx của bộ văn bản thử', () => {
  async function paragraphXml(prefix: string, needle: string): Promise<string> {
    const d = TEST_DOCS.find((x) => x.file.startsWith(prefix))!;
    const r = await formatNd30(d.lines.join('\n'));
    const xml = await (await JSZip.loadAsync(r.docx)).file('word/document.xml')!.async('string');
    const at = xml.indexOf(needle);
    expect(at, `không thấy "${needle}"`).toBeGreaterThan(-1);
    const start = xml.lastIndexOf('<w:p>', at);
    return xml.slice(start, xml.indexOf('</w:p>', at));
  }

  it('02: dòng "Theo đề nghị của …" in nghiêng như các dòng Căn cứ', async () => {
    expect(await paragraphXml('02', 'Theo đề nghị của')).toContain('<w:i/>');
    expect(await paragraphXml('02', 'Căn cứ Luật')).toContain('<w:i/>');
  });

  it('02: tên cơ quan ngắt thành hai dòng đậm, căn giữa', async () => {
    for (const line of ['ỦY BAN NHÂN DÂN<', 'THÀNH PHỐ MẪU']) {
      const p = await paragraphXml('02', line);
      expect(p).toContain('<w:b/>');
      expect(p).toContain('w:val="center"');
    }
  });

  it('03: "(Ban hành kèm theo …)" căn giữa, in nghiêng, không thụt đầu dòng', async () => {
    const p = await paragraphXml('03', '(Ban hành kèm theo');
    expect(p).toContain('<w:i/>');
    expect(p).toContain('w:val="center"');
    expect(p).not.toContain('w:firstLine');
  });
});
