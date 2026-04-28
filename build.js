// Local build helper for combining POS HTML fragments into index.html
// This file is intended to run with Node.js only, not in Google Apps Script.

const fs = require('fs');

function readFile(name) {
  return fs.readFileSync(name, 'utf8');
}

const html = [
  '<!DOCTYPE html>',
  '<html lang="th">',
  '<head>',
  '<meta charset="UTF-8">',
  '<meta name="viewport" content="width=device-width,initial-scale=1.0">',
  '<title>POS - ลาบบ้านสวน</title>',
  readFile('pos_styles.html'),
  '</head>',
  '<body>',
  readFile('pos_body.html'),
  readFile('pos_js1.html'),
  readFile('pos_js2.html'),
  '</body>',
  '</html>'
].join('\n');

fs.writeFileSync('index.html', html, 'utf8');
console.log('Built index.html OK -', fs.statSync('index.html').size, 'bytes');
