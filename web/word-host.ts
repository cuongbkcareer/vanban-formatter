/// <reference types="office-js" />
import { readFlatOpc } from '../src/docx-read';
import type { FixItem, FormatProfile } from '../src/learn/profile';
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

  async readFormat() {
    return Word.run(async (ctx) => {
      const ooxml = ctx.document.body.getOoxml();
      await ctx.sync();
      return readFlatOpc(ooxml.value);
    });
  },

  async selectParagraph(index) {
    await Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs;
      paras.load('items');
      await ctx.sync();
      paras.items[index]?.select();
      await ctx.sync();
    });
  },

  async applyFixes(items: FixItem[], _profile: FormatProfile) {
    const norm = (s: string): string => cleanWordText(s).replace(/\s+/g, ' ').trim();
    const ALIGN = { left: 'Left', center: 'Centered', right: 'Right', justify: 'Justified' } as const;
    const MM = 72 / 25.4;
    return Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs;
      paras.load('items/text');
      await ctx.sync();
      let applied = 0;
      let skipped = 0;
      let capsNotApplied = 0;
      for (const [n, it] of items.entries()) {
        const p = paras.items[it.index];
        // Chỉ sửa khi đúng đoạn đã phân tích (tránh sửa nhầm nếu tài liệu vừa thay đổi).
        if (!p || !norm(p.text).startsWith(norm(it.text))) {
          skipped++;
          continue;
        }
        const s = it.set;
        if (s.align) p.alignment = ALIGN[s.align];
        if (s.firstLineMm !== undefined) p.firstLineIndent = s.firstLineMm * MM;
        if (s.leftMm !== undefined) p.leftIndent = s.leftMm * MM;
        if (s.spaceBeforePt !== undefined) p.spaceBefore = s.spaceBeforePt;
        if (s.spaceAfterPt !== undefined) p.spaceAfter = s.spaceAfterPt;
        if (s.lineSpacing !== undefined) p.lineSpacing = s.lineSpacing * 12; // Office.js: điểm; Word chia cho 12
        if (s.sizePt !== undefined) p.font.size = s.sizePt;
        if (s.bold !== undefined) p.font.bold = s.bold;
        if (s.italic !== undefined) p.font.italic = s.italic;
        if (s.font) p.font.name = s.font;
        if (s.caps) capsNotApplied++;
        applied++;
        if (n % 40 === 39) await ctx.sync();
      }
      await ctx.sync();
      return { applied, skipped, capsNotApplied };
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
