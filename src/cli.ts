#!/usr/bin/env node
/**
 * CLI chạy hoàn toàn offline:
 *   npm run cli -- fixtures/cong-van.txt -o out/cong-van.docx
 *   npm run cli -- input.txt --preview            (chỉ xem nhãn từng dòng)
 *   npm run cli -- input.txt --issuer "SỞ X" --place "Hà Nội" -o out.docx
 */
import { readFile, writeFile } from 'node:fs/promises';
import { parseArgs } from 'node:util';
import { assemble, buildIR, classify, LABEL_NAMES, type Options } from './profiles/vn-nd30';
import { renderDocx } from './render-docx';

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    out: { type: 'string', short: 'o' },
    preview: { type: 'boolean', default: false },
    parent: { type: 'string' },
    issuer: { type: 'string' },
    place: { type: 'string' },
    type: { type: 'string' },
    date: { type: 'string' },
    number: { type: 'string' },
    'signer-title': { type: 'string' },
    'signer-name': { type: 'string' },
  },
});

const input = positionals[0];
if (!input) {
  console.error('Cách dùng: cli <file.txt> [-o out.docx] [--preview] [--issuer ..] [--parent ..] [--place ..] [--type ..] [--date YYYY-MM-DD]');
  process.exit(2);
}

const options: Options = {};
if (values.parent) options.parentOrg = values.parent;
if (values.issuer) options.issuer = values.issuer;
if (values.place) options.place = values.place;
if (values.type) options.docType = values.type;
if (values.date) options.dateIso = values.date;
if (values.number) options.number = values.number;
if (values['signer-title']) options.signerTitle = values['signer-title'].replace(/\\n/g, '\n');
if (values['signer-name']) options.signerName = values['signer-name'];

const raw = await readFile(input, 'utf8');
const lines = classify(raw);
const model = assemble(lines, options);

if (values.preview) {
  for (const l of lines) {
    const flag = l.confidence < 0.7 ? ' ⚠' : '';
    console.log(`[${LABEL_NAMES[l.label].padEnd(28)}] ${l.confidence.toFixed(2)}${flag}  ${l.text}`);
  }
}
for (const w of model.warnings) console.error(`Cảnh báo: ${w}`);

if (!values.preview || values.out) {
  const out = values.out ?? input.replace(/\.[^.]+$/, '') + '.docx';
  await writeFile(out, await renderDocx(buildIR(model)));
  console.error(`Đã tạo ${out}`);
}
