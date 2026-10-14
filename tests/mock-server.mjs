// Runs Code.gs (through the in-memory mock) behind a local HTTP endpoint so the
// front end can be tried end to end:  node tests/mock-server.mjs [port]
// then start the app with VITE_SHEETS_API_URL=http://localhost:8787/exec npm run dev
import http from 'node:http';
import { createGas } from './gas-mock.mjs';

const port = Number(process.argv[2] || 8787);
let gas = createGas();
let delay = 0;

const server = http.createServer((req, res) => {
  const cors = { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' };
  const url = new URL(req.url, 'http://x');
  if (url.pathname === '/__sheet') {
    // Test hook: inspect or edit the in-memory sheet.
    let body = '';
    req.on('data', (c) => (body += c));
    req.on('end', () => {
      if (req.method === 'POST') {
        const cmd = JSON.parse(body);
        if (cmd.reset) gas = createGas();
        if (cmd.appendRow) gas.sheet(cmd.sheet).data.push(cmd.appendRow);
        if (cmd.setRaw) gas.sheet(cmd.sheet).setRaw(cmd.setRaw);
        if ('delay' in cmd) delay = cmd.delay;
      }
      const out = {};
      for (const [name, sh] of gas.sheets) out[name] = sh.objects();
      res.writeHead(200, cors).end(JSON.stringify(out));
    });
    return;
  }
  if (url.pathname !== '/exec') return res.writeHead(404).end();
  let body = '';
  req.on('data', (c) => (body += c));
  req.on('end', () => {
    setTimeout(() => {
      const out = req.method === 'POST' ? gas.post(JSON.parse(body || '{}')) : gas.get();
      res.writeHead(200, cors).end(JSON.stringify(out));
    }, delay);
  });
});
server.listen(port, () => console.log(`mock sheet on http://localhost:${port}/exec`));
