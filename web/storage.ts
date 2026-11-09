/** Lưu cục bộ trong trình duyệt/khung add-in (localStorage). Không gửi đi đâu. Mọi lỗi lưu đều bị bỏ qua êm. */
import type { FormatProfile } from '../src/learn/profile';

const K_PROFILES = 'vanban.profiles.v1';
const K_ACTIVE = 'vanban.activeProfile.v1';
const K_OPTIONS = 'vanban.options.v1';

function read<T>(key: string, fallback: T): T {
  try {
    const v = localStorage.getItem(key);
    return v ? (JSON.parse(v) as T) : fallback;
  } catch {
    return fallback;
  }
}
function write(key: string, value: unknown): void {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    /* hết dung lượng hoặc bị chặn: bỏ qua */
  }
}

export function isProfile(x: unknown): x is FormatProfile {
  const p = x as FormatProfile | null;
  return (
    !!p &&
    p.version === 1 &&
    typeof p.name === 'string' &&
    Array.isArray(p.roles) &&
    p.roles.length > 0 &&
    p.roles.every((r) => r && typeof r.id === 'string' && r.format && typeof r.format.sizePt === 'number' && r.features)
  );
}

export const loadProfiles = (): FormatProfile[] => read<unknown[]>(K_PROFILES, []).filter(isProfile);
export const saveProfiles = (list: FormatProfile[]): void => write(K_PROFILES, list);
export const getActiveName = (): string => read<string>(K_ACTIVE, '');
export const setActiveName = (name: string): void => write(K_ACTIVE, name);

/** Thông tin cơ quan hay dùng (loại văn bản, cơ quan, địa danh, người ký ...). */
export const loadOptions = (): Record<string, string> => read<Record<string, string>>(K_OPTIONS, {});
export const saveOptions = (o: Record<string, string>): void => write(K_OPTIONS, o);
