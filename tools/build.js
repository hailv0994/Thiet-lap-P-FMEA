#!/usr/bin/env node
// Build file độc lập P-FMEA-Builder.html từ index.html + styles.css + js/*.js + data/*.js + vendor/*
// Chạy: node tools/build.js   (từ thư mục gốc của repo)
const fs = require('fs');
const path = require('path');

// Thư mục gốc repo = thư mục cha của thư mục chứa file build này
const ROOT = path.resolve(__dirname, '..');
const b64 = f => fs.readFileSync(path.join(ROOT, f)).toString('base64');
const css = fs.readFileSync(path.join(ROOT, 'styles.css'), 'utf8');

// Thứ tự nạp quan trọng
const order = [
  'vendor/xlsx.full.min.js', 'vendor/fflate.min.js',
  'data/severity.js', 'data/template.js', 'data/material.js',
  'js/parser.js', 'js/export-template.js', 'js/app.js'
];
const MODS = {}; order.forEach(f => MODS[f] = b64(f));

let html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
html = html.replace(/<link rel="stylesheet" href="styles.css" \/>/, "<style>\n" + css + "\n</style>");
// Gỡ <script src=...> NỘI BỘ; giữ lại script từ CDN (http/https) như Supabase
html = html.replace(/\s*<script src="(?!https?:)[^"]+"><\/script>/g, '');

const boot = '<scr' + 'ipt>\n(function(){\n'
  + 'function dec(b){var s=atob(b),u=new Uint8Array(s.length);for(var i=0;i<s.length;i++)u[i]=s.charCodeAt(i);return new TextDecoder("utf-8").decode(u);}\n'
  + 'var M=' + JSON.stringify(MODS) + ';\n'
  + '[' + order.map(f => JSON.stringify(f)).join(',') + '].forEach(function(k){(0,eval)(dec(M[k]));});\n'
  + '})();\n</scr' + 'ipt>';
html = html.replace('</body>', boot + '\n</body>');

fs.writeFileSync(path.join(ROOT, 'P-FMEA-Builder.html'), html);
console.log('Built standalone:', (fs.statSync(path.join(ROOT, 'P-FMEA-Builder.html')).size / 1024 / 1024).toFixed(2) + ' MB');
