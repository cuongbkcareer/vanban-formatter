/// <reference types="office-js" />
import type { Host } from './ui';

/** Bỏ ký tự điều khiển của Word (hết ô bảng \u0007, ngắt dòng mềm \u000b …). */
export function cleanWordText(s: string): string {
  return s
    .replace(/[\u000b\u000c]/g, '\n')
    .replace(/\r/g, '\n')
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u0008\u000e-\u001f]/g, '');
}

/** Cầu nối Office.js cho Word. Cần WordApi 1.3 trở lên (Windows, Mac, Word trên web). */
export const wordHost: Host = {
  kind: 'word',

  async readSource() {
    return Word.run(async (ctx) => {
      const selection = ctx.document.getSelection();
      selection.load('text');
      const selParas = selection.paragraphs;
      selParas.load('items/text');
      const bodyParas = ctx.document.body.paragraphs;
      bodyParas.load('items/text');
      await ctx.sync();
      const items = selection.text.trim().length > 0 ? selParas.items : bodyParas.items;
      return items.map((p) => cleanWordText(p.text)).join('\n');
    });
  },

  async openNew(base64) {
    await Word.run(async (ctx) => {
      const doc = ctx.application.createDocument(base64);
      doc.open();
      await ctx.sync();
    });
  },

  async replaceCurrent(base64) {
    await Word.run(async (ctx) => {
      const body = ctx.document.body;
      // Tham số tuỳ chọn (WordApi 1.5) chưa có trong @types/office-js ổn định.
      const insert = body.insertFileFromBase64.bind(body) as (
        base64File: string,
        insertLocation: 'Replace',
        options?: Record<string, boolean>,
      ) => unknown;
      if (Office.context.requirements.isSetSupported('WordApi', '1.5')) {
        insert(base64, 'Replace', {
          importTheme: false,
          importStyles: false,
          importParagraphSpacing: true,
          importPageColor: false,
          importDifferentOddEvenPages: false,
        });
      } else {
        insert(base64, 'Replace');
      }
      await ctx.sync();
    });
  },
};
