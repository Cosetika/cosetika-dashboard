const http = require('http');
const fs = require('fs');
const path = require('path');
const https = require('https');

const PORT = process.env.PORT || 3000;
const API_KEY = process.env.CONTIFICO_API_KEY || '';

let cache = { documentos: [], ultima_sync: null, sincronizando: false };

function fetchContifico(url) {
  return new Promise((resolve, reject) => {
    const req = https.request(url, {
      method: 'GET',
      headers: { 'Authorization': API_KEY, 'Accept': 'application/json' }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    req.setTimeout(15000, () => { req.destroy(); reject(new Error('Timeout')); });
    req.end();
  });
}

async function sincronizarHoy() {
  if (cache.sincronizando) return;
  cache.sincronizando = true;
  try {
    const now = new Date();
    const d = n => String(n).padStart(2, '0');
    const y = String(now.getFullYear()).slice(-2);
    const fecha = `${d(now.getDate())}/${d(now.getMonth()+1)}/${y}`;
    const url = `https://api.contifico.com/sistema/api/v1/documento/?tipo_documento=FAC&fecha_inicio=${fecha}&fecha_fin=${fecha}`;
    console.log('Sincronizando hoy:', fecha);
    const inicio = Date.now();
    const r = await fetchContifico(url);
    const tiempo = Date.now() - inicio;
    console.log(`Respuesta en ${tiempo}ms, status: ${r.status}`);
    const data = JSON.parse(r.body);
    const docs = Array.isArray(data) ? data : (data.results || data.data || []);
    console.log(`Facturas de hoy: ${docs.length}`);
    cache.documentos = docs;
    cache.ultima_sync = new Date().toISOString();
    cache.sincronizando = false;
    return { tiempo_ms: tiempo, facturas: docs.length, fecha };
  } catch(e) {
    console.error('Error sync:', e.message);
    cache.sincronizando = false;
    throw e;
  }
}

// Sincronizar al arrancar y cada hora
sincronizarHoy().catch(e => console.error('Error inicial:', e.message));
setInterval(() => sincronizarHoy().catch(e => console.error('Error sync:', e.message)), 60 * 60 * 1000);

const MIME = { '.html':'text/html', '.js':'application/javascript', '.css':'text/css', '.json':'application/json' };

const server = http.createServer(async (req, res) => {
  const urlPath = req.url.split('?')[0];
  res.setHeader('Access-Control-Allow-Origin', '*');
  if (req.method === 'OPTIONS') { res.writeHead(200); res.end(); return; }

  if (urlPath === '/api/hoy') {
    try {
      const inicio = Date.now();
      const result = await sincronizarHoy();
      res.writeHead(200, {'Content-Type': 'application/json'});
      res.end(JSON.stringify(result));
    } catch(e) {
      res.writeHead(200, {'Content-Type': 'application/json'});
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  if (urlPath === '/api/debug') {
    try {
      const inicio = Date.now();
      const r = await fetchContifico('https://api.contifico.com/sistema/api/v1/marca/');
      const tiempo = Date.now() - inicio;
      res.writeHead(200, {'Content-Type': 'application/json'});
      res.end(JSON.stringify({ 
        key: API_KEY.substring(0,8)+'...', 
        status: r.status, 
        tiempo_ms: tiempo,
        marcas: r.body.substring(0,300),
        cache_docs: cache.documentos.length,
        ultima_sync: cache.ultima_sync
      }));
    } catch(e) {
      res.writeHead(200, {'Content-Type': 'application/json'});
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  if (urlPath === '/api/ventas') {
    res.writeHead(200, {'Content-Type': 'application/json'});
    res.end(JSON.stringify({
      total: cache.documentos.length,
      documentos: cache.documentos,
      ultima_sync: cache.ultima_sync
    }));
    return;
  }

  // Static files
  let filePath = urlPath === '/' ? path.join(__dirname, 'index.html')
    : urlPath === '/login' ? path.join(__dirname, 'public', 'login.html')
    : path.join(__dirname, urlPath);

  if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
    res.writeHead(200, {'Content-Type': MIME[path.extname(filePath)] || 'text/plain'});
    fs.createReadStream(filePath).pipe(res);
  } else {
    res.writeHead(200, {'Content-Type': 'text/html'});
    fs.createReadStream(path.join(__dirname, 'index.html')).pipe(res);
  }
});

server.listen(PORT, '0.0.0.0', () => console.log(`Cosétika Dashboard running on port ${PORT}`));
