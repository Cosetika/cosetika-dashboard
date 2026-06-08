const http = require('http');
const fs = require('fs');
const path = require('path');
const https = require('https');

const PORT = process.env.PORT || 3000;
const API_KEY = process.env.CONTIFICO_API_KEY || '';

function fetchContifico(url) {
  return new Promise((resolve, reject) => {
    const options = {
      method: 'GET',
      headers: { 'Authorization': API_KEY, 'Accept': 'application/json' }
    };
    const req = https.request(url, options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    req.setTimeout(25000, () => { req.destroy(); reject(new Error('Timeout')); });
    req.end();
  });
}

const MIME = {
  '.html': 'text/html',
  '.js': 'application/javascript', 
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.ico': 'image/x-icon'
};

const server = http.createServer(async (req, res) => {
  const urlPath = req.url.split('?')[0];
  
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') { res.writeHead(200); res.end(); return; }

  // API routes
  if (urlPath === '/api/ventas') {
    if (!API_KEY) {
      res.writeHead(500, {'Content-Type': 'application/json'});
      res.end(JSON.stringify({ error: 'API Key no configurada' }));
      return;
    }
    try {
      const r = await fetchContifico(
        'https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC&offset=0&limit=100'
      );
      let data;
      try { data = JSON.parse(r.body); } catch(e) {
        res.writeHead(500, {'Content-Type': 'application/json'});
        res.end(JSON.stringify({ error: 'JSON inválido', raw: r.body.substring(0, 200) }));
        return;
      }
      const docs = Array.isArray(data) ? data : (data.results || data.data || []);
      const procesados = docs.map(doc => ({
        fecha: doc.fecha_emision || '',
        vendedor: doc.vendedor_nombre || doc.vendedor || 'Sin vendedor',
        cliente: doc.cliente_razon_social || doc.razon_social || '',
        provincia: doc.provincia || '',
        total: parseFloat((doc.total || '0').toString().replace(',', '.')),
        estado: doc.estado || '',
        detalles: (doc.detalles || []).map(d => ({
          producto: d.producto_nombre || d.nombre || '',
          marca: d.adicional3 || d.marca || '',
          cantidad: parseFloat((d.cantidad || '0').toString().replace(',', '.')),
          total: parseFloat((d.base_gravable || '0').toString().replace(',', '.'))
        }))
      }));
      res.writeHead(200, {'Content-Type': 'application/json'});
      res.end(JSON.stringify({ total: procesados.length, documentos: procesados }));
    } catch(e) {
      res.writeHead(500, {'Content-Type': 'application/json'});
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  if (urlPath === '/api/debug') {
    try {
      const r = await fetchContifico('https://api.contifico.com/sistema/api/v1/marca/');
      res.writeHead(200, {'Content-Type': 'application/json'});
      res.end(JSON.stringify({ key: API_KEY.substring(0,8)+'...', status: r.status, resp: r.body.substring(0, 300) }));
    } catch(e) {
      res.writeHead(200, {'Content-Type': 'application/json'});
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // Static files
  let filePath;
  if (urlPath === '/' || urlPath === '') {
    filePath = path.join(__dirname, 'index.html');
  } else if (urlPath === '/login') {
    filePath = path.join(__dirname, 'public', 'login.html');
  } else {
    filePath = path.join(__dirname, urlPath);
  }

  if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
    const ext = path.extname(filePath);
    res.writeHead(200, {'Content-Type': MIME[ext] || 'text/plain'});
    fs.createReadStream(filePath).pipe(res);
  } else {
    // SPA fallback - serve index.html
    const index = path.join(__dirname, 'index.html');
    if (fs.existsSync(index)) {
      res.writeHead(200, {'Content-Type': 'text/html'});
      fs.createReadStream(index).pipe(res);
    } else {
      res.writeHead(404);
      res.end('Not found');
    }
  }
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Cosétika Dashboard running on port ${PORT}`);
});
