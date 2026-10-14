#!/usr/bin/env node
/* 把 src/ui.html、src/core.js、內建範本與第三方函式庫打包成單一離線 HTML（dist/）。
   用法：npm run build */
'use strict';
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const root = path.resolve(__dirname, '..');
const read = p => fs.readFileSync(path.join(root, p), 'utf8');
const lib = p => fs.readFileSync(require.resolve(p), 'utf8');

// 內嵌在 <script> 裡時，避免提早結束標籤；拿掉 source map 參照（離線用不到）
const inline = js => js.replace(/<\/(script)/gi, '<\\/$1').replace(/\/\/# sourceMappingURL=.*$/gm, '');

let html = read('src/ui.html');
const parts = {
  __PDFJS__: lib('pdfjs-dist/build/pdf.min.js'),
  __PDFWORKER__: lib('pdfjs-dist/build/pdf.worker.min.js'), // 主執行緒模式，不另開 worker、不需網路
  __JSZIP__: lib('jszip/dist/jszip.min.js'),
  __CORE__: read('src/core.js'),
};
for (const [key, js] of Object.entries(parts)) {
  const tag = `<script>/*${key}*/</script>`;
  if (html.split(tag).length !== 2) throw new Error(`ui.html 找不到唯一的 ${tag}`);
  html = html.split(tag).join(`<script>\n${inline(js)}\n</script>`);
}
// 內建範本：作者範本、通用範本
for (const [key, file] of [['__AUTHOR_B64__', 'template/author.xlsx'], ['__GENERIC_B64__', 'template/generic.xlsx']]) {
  if (!html.includes(key)) throw new Error(`ui.html 找不到 ${key}`);
  html = html.replace(key, fs.readFileSync(path.join(root, file)).toString('base64'));
}

const out = path.join(root, 'dist', '謄本轉地號清冊.html');
fs.mkdirSync(path.dirname(out), { recursive: true });
fs.writeFileSync(out, html);
const sha = crypto.createHash('sha256').update(html).digest('hex');
console.log(`✓ ${path.relative(root, out)}  ${(Buffer.byteLength(html) / 1e6).toFixed(2)} MB`);
console.log(`  SHA-256 ${sha}`);
