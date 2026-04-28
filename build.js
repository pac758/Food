const fs = require('fs');
const s = fs.readFileSync('pos_styles.html', 'utf8');
const b = fs.readFileSync('pos_body.html', 'utf8');
const j1 = fs.readFileSync('pos_js1.html', 'utf8');
const j2 = fs.readFileSync('pos_js2.html', 'utf8');
const html = [
  '<!DOCTYPE html>',
  '<html lang="th">',
  '<head>',
  '<meta charset="UTF-8">',
  '<meta name="viewport" content="width=device-width,initial-scale=1.0">',
  '<title>POS - ลาบบ้านสวน</title>',
  s,
  '</head>',
  '<body>',
  b,
  j1,
  j2,
  '</body>',
  '</html>'
].join('\n');
fs.writeFileSync('index.html', html, 'utf8');
console.log('Built index.html OK -', fs.statSync('index.html').size, 'bytes');
