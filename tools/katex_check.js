// Đọc JSON mảng chuỗi LaTeX từ stdin, in JSON mảng (null = hợp lệ, chuỗi = thông báo lỗi).
const katex = require('katex');
let buf = '';
process.stdin.on('data', d => (buf += d));
process.stdin.on('end', () => {
  const items = JSON.parse(buf);
  const out = items.map(s => {
    try { katex.renderToString(s, { throwOnError: true, strict: 'error' }); return null; }
    catch (e) { return String(e.message).slice(0, 200); }
  });
  process.stdout.write(JSON.stringify(out));
});
