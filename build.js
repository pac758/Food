// Build script: Extract component files from index.html (source of truth)
// Usage: node build.js extract  — splits index.html into component files
// Usage: node build.js build    — combines component files into index.html

const fs = require('fs');
const path = require('path');

const dir = __dirname;
const indexPath = path.join(dir, 'index.html');

function extract() {
  const html = fs.readFileSync(indexPath, 'utf8');
  
  // Extract <style>...</style> + <link> (first occurrence)
  const styleMatch = html.match(/<style>([\s\S]*?)<\/style>\s*(<link[^>]*>)?/);
  if (styleMatch) {
    const css = '<style>\n' + styleMatch[1].trim() + '\n</style>\n' + (styleMatch[2] || '') + '\n';
    fs.writeFileSync(path.join(dir, 'pos_styles.html'), css, 'utf8');
    console.log('✅ pos_styles.html extracted');
  }

  // Extract body content (between </head><body> and first <script>)
  const bodyMatch = html.match(/<body>([\s\S]*?)(?=<script>)/);
  if (bodyMatch) {
    fs.writeFileSync(path.join(dir, 'pos_body.html'), bodyMatch[1].trim() + '\n', 'utf8');
    console.log('✅ pos_body.html extracted');
  }

  // Extract script blocks
  const scripts = [];
  const scriptRegex = /<script>([\s\S]*?)<\/script>/g;
  let m;
  while ((m = scriptRegex.exec(html)) !== null) {
    scripts.push(m[0]);
  }
  if (scripts.length >= 1) {
    fs.writeFileSync(path.join(dir, 'pos_js1.html'), scripts[0] + '\n', 'utf8');
    console.log('✅ pos_js1.html extracted');
  }
  if (scripts.length >= 2) {
    fs.writeFileSync(path.join(dir, 'pos_js2.html'), scripts[1] + '\n', 'utf8');
    console.log('✅ pos_js2.html extracted');
  }
  console.log('🎉 Done! Component files synced from index.html');
}

function build() {
  const styles = fs.readFileSync(path.join(dir, 'pos_styles.html'), 'utf8').trim();
  const body = fs.readFileSync(path.join(dir, 'pos_body.html'), 'utf8').trim();
  const js1 = fs.readFileSync(path.join(dir, 'pos_js1.html'), 'utf8').trim();
  const js2 = fs.readFileSync(path.join(dir, 'pos_js2.html'), 'utf8').trim();

  const html = `<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>POS - ลาบบ้านสวน</title>
${styles}

</head>
<body>
${body}

${js1}

${js2}

</body>
</html>
`;
  fs.writeFileSync(indexPath, html, 'utf8');
  console.log('✅ index.html built from components');
}

const cmd = process.argv[2] || 'extract';
if (cmd === 'build') build();
else extract();
