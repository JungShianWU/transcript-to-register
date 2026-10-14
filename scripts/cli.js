#!/usr/bin/env node
/* 命令列版本：不開瀏覽器，直接把謄本 PDF 轉成清冊。
   用法：
     node scripts/cli.js -o 輸出.xlsx [-t author|generic|範本.xlsx] [--author 名字] 謄本1.pdf 謄本2.pdf ...
   -t author：作者範本（預設）；-t generic：通用範本；或指定自己的地號清冊範本檔。 */
'use strict';
const fs = require('fs');
const path = require('path');
const pdfjs = require('pdfjs-dist/legacy/build/pdf.js');
const JSZip = require('jszip');
require('../src/core.js');
const C = globalThis.TranscriptCore;

function args(argv) {
  const a = { pdfs: [], author: 'Jacky' };
  for (let i = 0; i < argv.length; i++) {
    const v = argv[i];
    if (v === '-o') a.out = argv[++i];
    else if (v === '-t') a.template = argv[++i];
    else if (v === '--author') a.author = argv[++i];
    else a.pdfs.push(v);
  }
  if (!a.out || !a.pdfs.length) {
    console.error('用法：node scripts/cli.js -o 輸出.xlsx [-t author|generic|範本.xlsx] [--author 名字] 謄本.pdf ...');
    process.exit(2);
  }
  return a;
}

(async () => {
  const a = args(process.argv.slice(2));
  const items = [], warn = [];
  for (const f of a.pdfs) {
    const meta = {};
    const text = await C.pdfText(pdfjs, new Uint8Array(fs.readFileSync(f)), meta);
    const r = C.parseText(text, path.basename(f), meta);
    items.push(...r.items); warn.push(...r.warn);
  }
  const model = C.buildModel(items);
  if (!model.blocks.length) throw new Error(warn.length ? warn.join('\n') : '沒有讀到任何土地謄本');
  const mode = a.template || 'author';
  const tplPath = mode === 'author' || mode === 'generic' ? path.join(__dirname, '..', 'template', mode + '.xlsx') : mode;
  const fill = mode === 'author' ? C.fillAuthorWorkbook : C.fillWorkbook;
  const res = await fill(JSZip, fs.readFileSync(tplPath), model, { author: a.author });
  if (mode === 'author') model.info = (model.info || []).filter(s => !/隱藏欄/.test(s)).concat(res.info || []);
  fs.writeFileSync(a.out, res.buf);
  console.log(`✓ ${a.out}：土地 ${model.lands.length} 筆，寫入 ${res.rows} 列`);
  (model.info || []).forEach(s => console.log('  ℹ ' + s));
  [...warn, ...model.warn].forEach(s => console.log('  ⚠ ' + s));
})().catch(e => { console.error(e.message || e); process.exit(1); });
