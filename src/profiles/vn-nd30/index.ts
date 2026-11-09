import { renderDocx } from '../../render-docx';
import { assemble } from './assemble';
import { buildIR } from './build-ir';
import { classify } from './classify';
import type { LabeledLine, Model, Options } from './types';

export { assemble } from './assemble';
export { buildIR, ND30_LAYOUT, formatDate } from './build-ir';
export { classify, parsePlaceDate } from './classify';
export * from './types';

export interface FormatResult {
  docx: Uint8Array;
  lines: LabeledLine[];
  model: Model;
}

/**
 * Toàn bộ quy trình: văn bản thô → nhãn → mô hình → IR → .docx.
 * Nếu truyền `lines` (đã được người dùng sửa nhãn) thì bỏ qua bước phân loại.
 */
export async function formatNd30(
  raw: string,
  options: Options = {},
  lines?: LabeledLine[],
): Promise<FormatResult> {
  const labeled = lines ?? classify(raw);
  const model = assemble(labeled, options);
  const docx = await renderDocx(buildIR(model));
  return { docx, lines: labeled, model };
}
