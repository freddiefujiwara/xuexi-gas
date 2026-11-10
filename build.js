// src/main.js から Apps Script 用の dist/Code.gs を生成する。
// export 文は正規表現で除去するため、src 側の export は必ず1行で書くこと。
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';

const src = readFileSync(new URL('./src/main.js', import.meta.url), 'utf8');
const code = src.replace(/^export\s*\{[^}]*\};?\s*$/m, '');
mkdirSync(new URL('./dist', import.meta.url), { recursive: true });
writeFileSync(new URL('./dist/Code.gs', import.meta.url), code);
console.log('dist/Code.gs written', code.length, 'chars');
