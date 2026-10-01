// 미리보기 HTML 생성: src/Index.html 에 가짜 서버(mock-server.js)를 주입한다.
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
const html = readFileSync('src/Index.html', 'utf8');
const mock = readFileSync('tools/mock-server.js', 'utf8');
const out = process.argv[2] || 'preview/index.html';
mkdirSync(out.replace(/\/[^/]*$/, ''), { recursive: true });
writeFileSync(out, html.replace('<script>', `<script>\n${mock}\n</script>\n<script>`));
console.log('wrote', out);
