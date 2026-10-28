const fs = require('fs');
const path = require('path');

const outPath = process.argv[2];
if (!outPath) {
  console.error('Usage: node scripts/package-elements.js <output-file>');
  process.exit(1);
}

const parts = [
  'dist/generated-daxe-xl/runtime.js',
  'dist/generated-daxe-xl/polyfills.js',
  'dist/generated-daxe-xl/main.js'
].map(f => fs.readFileSync(f));

fs.mkdirSync(path.dirname(outPath), { recursive: true });
fs.writeFileSync(outPath, Buffer.concat(parts));
console.log('Wrote', outPath);
