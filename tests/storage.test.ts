import { describe, expect, it } from 'vitest';
import { isProfile } from '../web/storage';

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
