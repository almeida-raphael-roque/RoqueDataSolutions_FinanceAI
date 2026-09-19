import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;

app.use(express.static(path.join(__dirname, 'gas-project')));
app.use('/assets', express.static(path.join(__dirname, 'assets')));

function renderIndex(req, res) {
  let html = fs.readFileSync(path.join(__dirname, 'gas-project', 'Index.html'), 'utf8');
  html = html.replace(/<\?\!= include\('(.*?)'\); \?>/g, (match, p1) => {
    let p1Name = p1.endsWith('.html') ? p1 : p1 + '.html';
    try {
      return fs.readFileSync(path.join(__dirname, 'gas-project', p1Name), 'utf8');
    } catch(e) {
      return `<!-- Error including ${p1Name} -->`;
    }
  });
  res.send(html);
}

app.get('/', renderIndex);
app.get('/index.html', renderIndex);

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server running on port ${PORT}`);
});
