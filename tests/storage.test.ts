import { describe, expect, it } from 'vitest';
import { isProfile, pickOrgFields } from '../web/storage';

describe('hồ sơ mẫu định dạng (nhập từ file .json)', () => {
  const ok = {
    version: 1,
    name: 'Mẫu A',
    roles: [{ id: 'r1', name: 'Nội dung', count: 3, format: { sizePt: 13 }, features: {} }],
  };
  it('chấp nhận hồ sơ đúng cấu trúc', () => {
    expect(isProfile(ok)).toBe(true);
  });
  it('từ chối dữ liệu sai hoặc thiếu', () => {
    expect(isProfile(null)).toBe(false);
    expect(isProfile({ ...ok, version: 2 })).toBe(false);
    expect(isProfile({ ...ok, roles: [] })).toBe(false);
    expect(isProfile({ ...ok, roles: [{ id: 'r1' }] })).toBe(false);
    expect(isProfile('x')).toBe(false);
  });
});

describe('thông tin cơ quan được nhớ', () => {
  it('chỉ giữ các ô thuộc cơ quan, bỏ loại văn bản, số/ký hiệu và ngày', () => {
    const saved = pickOrgFields({
      'o-parent': 'ỦY BAN NHÂN DÂN TỈNH MẪU',
      'o-issuer': 'SỞ NỘI VỤ',
      'o-place': 'Hà Nội',
      'o-stitle': 'GIÁM ĐỐC',
      'o-sname': 'Nguyễn Văn A',
      'o-type': 'Quyết định',
      'o-number': '12/QĐ-UBND',
      'o-date': '2025-03-05',
    });
    expect(Object.keys(saved).sort()).toEqual(['o-issuer', 'o-parent', 'o-place', 'o-sname', 'o-stitle']);
  });
  it('bỏ ô trống và giá trị không phải chuỗi', () => {
    expect(pickOrgFields({ 'o-place': '  ', 'o-issuer': 5, 'o-sname': 'A' })).toEqual({ 'o-sname': 'A' });
  });
});
