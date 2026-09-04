import http from 'http';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = 3000;

function include(filename) {
  try {
    return fs.readFileSync(path.join(__dirname, 'gas-project', filename + '.html'), 'utf8');
  } catch (e) {
    return '';
  }
}

const server = http.createServer((req, res) => {
  if (req.url === '/' || req.url === '/index.html') {
    try {
      let html = fs.readFileSync(path.join(__dirname, 'gas-project', 'Index.html'), 'utf8');
      
      // Simulate GAS <?!= include('...') ?> tags
      html = html.replace(/<\?!=\s*include\('([^']+)'\);\s*\?>/g, (match, p1) => {
        return include(p1);
      });
      
      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end(html);
    } catch(e) {
      res.writeHead(500);
      res.end('Error loading Index.html');
    }
  } else {
    res.writeHead(404);
    res.end('Not found');
  }
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Preview server running on port ${PORT}`);
});
