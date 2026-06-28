const http = require('http');
const fs   = require('fs');
const path = require('path');

const PORT = 3000;
const DIR  = __dirname;

const mime = {
  '.html': 'text/html',
  '.mp4' : 'video/mp4',
  '.js'  : 'text/javascript',
  '.css' : 'text/css',
};

http.createServer((req, res) => {
  // decode URL so spaces and special chars work
  let filePath = path.join(DIR, decodeURIComponent(req.url === '/' ? '/index.html' : req.url));
  const ext    = path.extname(filePath).toLowerCase();
  const type   = mime[ext] || 'application/octet-stream';

  fs.stat(filePath, (err, stat) => {
    if (err) {
      res.writeHead(404);
      return res.end('Not found');
    }

    // ── Range support for video seeking ──
    const range = req.headers.range;
    if (range && ext === '.mp4') {
      const parts  = range.replace(/bytes=/, '').split('-');
      const start  = parseInt(parts[0], 10);
      const end    = parts[1] ? parseInt(parts[1], 10) : stat.size - 1;
      const chunk  = end - start + 1;

      res.writeHead(206, {
        'Content-Range' : `bytes ${start}-${end}/${stat.size}`,
        'Accept-Ranges' : 'bytes',
        'Content-Length': chunk,
        'Content-Type'  : type,
      });
      fs.createReadStream(filePath, { start, end }).pipe(res);
    } else {
      res.writeHead(200, { 'Content-Type': type, 'Content-Length': stat.size });
      fs.createReadStream(filePath).pipe(res);
    }
  });
}).listen(PORT, () => {
  console.log(`\n  ✓  Server running at  http://localhost:${PORT}\n`);
});
