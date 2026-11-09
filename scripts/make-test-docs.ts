// Tạo các file Word thô để thử add-in:  npx tsx scripts/make-test-docs.ts
// Kiểu chữ cố ý là mặc định của Word (không căn lề, không đậm) để giống văn bản thô.
import { Document, Packer, Paragraph, Tab, TextRun } from 'docx';
import { mkdir, writeFile } from 'node:fs/promises';
import { TEST_DOCS } from '../fixtures/test-docs';

const OUT = 'test-docs';
await mkdir(OUT, { recursive: true });

for (const d of TEST_DOCS) {
  const paragraphs = d.lines.map((line) => {
    const parts = line.split('\t');
    const children: (TextRun | Tab)[] = [];
    parts.forEach((p, i) => {
      if (i > 0) children.push(new TextRun({ children: [new Tab()] }));
      if (p) children.push(new TextRun({ text: p }));
    });
    return new Paragraph({ children });
  });
  const doc = new Document({
    styles: { default: { document: { run: { font: 'Calibri', size: 22 } } } },
    sections: [{ children: paragraphs }],
  });
  await writeFile(`${OUT}/${d.file}.docx`, await Packer.toBuffer(doc));
}

const readme = [
  'BỘ FILE WORD ĐỂ THỬ ADD-IN "CHUẨN HOÁ NĐ 30"',
  'Mở từng file trong Word, bấm nút "Chuẩn hoá NĐ 30" > "Chuẩn hoá văn bản này".',
  'Tất cả tên cơ quan, người, số hiệu đều là giả định.',
  '',
  ...TEST_DOCS.flatMap((d) => [`${d.file}.docx`, `  Thử: ${d.title}`, `  Kết quả mong đợi: ${d.expect}`, '']),
].join('\n');
await writeFile(`${OUT}/DOC-TRUOC-KHI-THU.txt`, readme);
console.log(`Đã tạo ${TEST_DOCS.length} file trong ${OUT}/`);
